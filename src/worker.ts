import { Hono } from "hono";
import { bodyLimit } from "hono/body-limit";
import { z } from "zod";
import { createAccess, NoShare, SHARE_ROLES, SUBJECT_KIND_NAMES, type ShareRole } from "./access.js";
import { authFromEnv, ConfigError, type Authenticator, type AuthEnv } from "./auth.js";
import { SheetError, type Principal } from "./core/workbook.js";
import { createMemoryBlobs, createMemoryStore, EraseRefusedError, type Blobs, type Store } from "agentic-identity";
import { message, OpError, OPS, removeWorkbook } from "./ops.js";
import { VERSION } from "./version.js";

export interface WorkerEnv extends AuthEnv {
  PACK_NAME?: string;
}

const MAX_BODY_BYTES = 5 * 1024 * 1024;
// One request may write this many cells. A cell costs memory in the isolate, which all callers share.
const MAX_WRITE_CELLS = 10_000;

// Each route runs one operation, the same one the CLI and the MCP server run.
// The path gives `book`. The query string and the JSON body give the rest.
const ROUTES = [
  ["get", "/", "list"],
  ["post", "/", "new"],
  ["get", "/:book/range", "read"],
  ["put", "/:book/range", "write"],
  ["put", "/:book/formula", "formula"],
  ["post", "/:book/rows", "insert-rows"],
  ["get", "/:book/export", "export"],
  ["post", "/:book/import", "import"],
] as const;

export interface AppOptions {
  store: Store;
  blobs: Blobs;
  authenticate: Authenticator;
  pack?: string;
}

export function createApp({ store, blobs, authenticate, pack = "agentic-spreadsheet" }: AppOptions): Hono {
  const app = new Hono();
  app.use(async (c, next) => {
    await next();
    c.header("cache-control", "no-store");
    c.header("x-content-type-options", "nosniff");
  });
  app.get("/health", (c) => c.json({ ok: true, pack, version: VERSION }));

  const access = createAccess(store);
  // Who may use a book: a principal with no role sees "no workbook", as for a missing book. `list` still names every book (see the README).
  async function guard(who: Principal, action: "read" | "write" | "admin", book: string): Promise<Response | null> {
    const { allowed, role } = await access.role(who, action, book);
    if (allowed) return null;
    return role === null ? Response.json({ error: `no workbook "${book}"` }, { status: 404 }) : Response.json({ error: `the ${role} role cannot ${action} this workbook` }, { status: 403 });
  }

  const api = new Hono<{ Variables: { who: Principal } }>();
  api.use(bodyLimit({ maxSize: MAX_BODY_BYTES, onError: (c) => c.json({ error: "the request is too large" }, 413) }));
  api.use(async (c, next) => {
    const who = await authenticate(c.req.raw);
    if (who === null) return c.json({ error: "sign in with a valid credential" }, 401);
    c.set("who", who);
    await next();
  });
  for (const [method, path, name] of ROUTES) {
    const op = OPS.find((o) => o.name === name)!;
    api.on(method, path, async (c) => {
      const book = c.req.param("book");
      const denied = book === undefined ? null : await guard(c.get("who"), op.readOnly ? "read" : "write", book);
      if (denied !== null) return denied;
      let body: unknown = {};
      if (method !== "get") {
        body = await c.req.json().catch(() => null);
        if (typeof body !== "object" || body === null || Array.isArray(body)) return c.json({ error: "the body must be a JSON object" }, 400);
      }
      const parsed = z.object(op.shape).strict().safeParse({ ...c.req.query(), ...(body as object), ...c.req.param() });
      if (!parsed.success) return c.json({ error: message(parsed.error) }, 400);
      const result = await op.run(parsed.data, { store, blobs, principal: () => c.get("who"), maxWriteCells: MAX_WRITE_CELLS });
      if (name === "new") await access.create(c.get("who"), (parsed.data as { book: string }).book);
      if (name === "export") return c.body((result as { csv: string }).csv, 200, { "content-type": "text/csv; charset=utf-8" });
      return c.json(result as object, name === "new" ? 201 : 200);
    });
  }
  // The owner shares a book with a principal by id, and takes the share back.
  const shareBody = z.object({ kind: z.enum(SUBJECT_KIND_NAMES), subject: z.string().regex(/^[\w.@:-]{1,128}$/), role: z.enum(SHARE_ROLES) }).strict();
  api.put("/:book/share", async (c) => {
    const denied = await guard(c.get("who"), "admin", c.req.param("book"));
    if (denied !== null) return denied;
    const parsed = shareBody.safeParse(await c.req.json().catch(() => null));
    if (!parsed.success) return c.json({ error: message(parsed.error) }, 400);
    await access.share(c.get("who"), c.req.param("book"), { kind: parsed.data.kind, id: parsed.data.subject }, parsed.data.role as ShareRole);
    return c.json({ ok: true });
  });
  api.delete("/:book/share/:kind/:subject", async (c) => {
    const kind = z.enum(SUBJECT_KIND_NAMES).safeParse(c.req.param("kind"));
    if (!kind.success) return c.json({ error: "the kind is human or agent" }, 400);
    const denied = await guard(c.get("who"), "admin", c.req.param("book"));
    if (denied !== null) return denied;
    await access.unshare(c.get("who"), c.req.param("book"), { kind: kind.data, id: c.req.param("subject") }).catch((error: unknown) => {
      if (error instanceof NoShare) throw new OpError(error.message, 404);
      throw error;
    });
    return c.json({ ok: true });
  });
  // The owner deletes a book. Its grants, invites, links, and audit entries go with it.
  api.delete("/:book", async (c) => {
    const denied = await guard(c.get("who"), "admin", c.req.param("book"));
    if (denied !== null) return denied;
    // The grants go after the blob and before the document, so a failed step can run again with the
    // owner's grant in place, and no one can make a book of this name while the old grants stand.
    await removeWorkbook({ store, blobs, principal: () => c.get("who") }, c.req.param("book"), () => access.erase(c.get("who"), c.req.param("book")));
    return c.json({ ok: true });
  });
  app.route("/api/v1/sheets", api);

  app.notFound((c) => c.json({ error: "not found" }, 404));
  // A message from an operation is for the caller. Anything else is a bug: log its name, not its text.
  app.onError((error, c) => {
    if (error instanceof EraseRefusedError) return c.json({ error: error.message }, 409);
    if (error instanceof OpError) return c.json({ error: error.message }, error.status);
    if (error instanceof SheetError || error.constructor === Error) return c.json({ error: error.message }, 400);
    console.error("sheets api error", error.name);
    return c.json({ error: "something failed on the server" }, 500);
  });
  return app;
}

const store = createMemoryStore();
const blobs = createMemoryBlobs();
let cached: { keys?: string; unsafe?: string; auth: Authenticator } | undefined;

// Data lives in memory in one isolate. It is lost when the isolate restarts.
export default {
  async fetch(request: Request, env: WorkerEnv): Promise<Response> {
    let authenticate = cached?.keys === env.SHEETS_API_KEYS && cached?.unsafe === env.SHEETS_UNSAFE_TRUST_HEADERS ? cached?.auth : undefined;
    try {
      authenticate ??= await authFromEnv(env);
      cached = { keys: env.SHEETS_API_KEYS, unsafe: env.SHEETS_UNSAFE_TRUST_HEADERS, auth: authenticate };
    } catch (error) {
      if (!(error instanceof ConfigError)) throw error;
      cached = undefined;
      return Response.json({ error: `the API cannot start: ${error.message}` }, { status: 500 });
    }
    return createApp({ store, blobs, authenticate, pack: env.PACK_NAME }).fetch(request);
  },
};
