import { Hono, type Context } from "hono";
import { bodyLimit } from "hono/body-limit";
import { z } from "zod";
import { createAccess, NoShare, SHARE_ROLES, SUBJECT_KIND_NAMES, type ShareRole } from "./access.js";
import { authFromEnv, ConfigError, type Authenticator, type AuthEnv } from "./auth.js";
import { SheetError, type Principal } from "./core/workbook.js";
import { EraseRefusedError, hasScope, parseRegion, type Auth, type Blobs, type PrincipalContext, type Role, type Region, type Store } from "agentic-identity";
import { principalFromWebSocketRequest } from "agentic-identity/http";
import { createD1Store, createR2Blobs, D1_SCHEMA, type D1Like, type R2Like } from "agentic-identity/cloudflare";
import { message, OpError, OPS, removeWorkbook } from "./ops.js";
import { VERSION } from "./version.js";

export { fromAuth } from "./auth.js";

export interface WorkerEnv extends AuthEnv {
  PACK_NAME?: string;
  /** Where the data lives: "eu" or "us". It is a claim you make at deploy time, with no default. */
  SHEETS_REGION?: string;
  /** Set to "1" to accept the region "unknown" under `wrangler dev`. Local D1 reports no region. Never set it in production. */
  SHEETS_UNSAFE_REGION?: string;
  /** D1 database: the workbooks and the shares. Create it with `wrangler d1 create agentic-spreadsheet --location weur`. */
  DB?: D1Like;
  /** R2 bucket: each workbook as a Yjs update. Create it with `--jurisdiction eu`. */
  BLOBS?: R2Like;
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
  ["delete", "/:book/rows", "delete-rows"],
  ["post", "/:book/cols", "insert-cols"],
  ["delete", "/:book/cols", "delete-cols"],
  ["get", "/:book/sheets", "sheets"],
  ["post", "/:book/sheets", "add-sheet"],
  ["delete", "/:book/sheets/:name", "delete-sheet"],
  ["get", "/:book/export", "export"],
  ["post", "/:book/import", "import"],
] as const;

export interface AppOptions {
  store: Store;
  blobs: Blobs;
  authenticate: Authenticator;
  pack?: string;
}

type ApiEnv = { Variables: { who: Principal; scopes: string[] | undefined } };

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
  // A credential may carry scopes that limit it below its role: a key with only `workbook:read` cannot write, whatever its role.
  async function guard(c: Context<ApiEnv>, action: "read" | "write" | "admin", book: string): Promise<Response | null> {
    const scopes = c.get("scopes");
    const { allowed, role } = await access.role(c.get("who"), action, book);
    if (allowed && role !== null && scopes !== undefined && !hasScope(scopes, `workbook:${action}`)) {
      return Response.json({ error: `this credential has no ${action} scope` }, { status: 403 });
    }
    if (allowed) return null;
    return role === null ? Response.json({ error: `no workbook "${book}"` }, { status: 404 }) : Response.json({ error: `the ${role} role cannot ${action} this workbook` }, { status: 403 });
  }

  const api = new Hono<ApiEnv>();
  api.use(bodyLimit({ maxSize: MAX_BODY_BYTES, onError: (c) => c.json({ error: "the request is too large" }, 413) }));
  api.use(async (c, next) => {
    const caller = await authenticate(c.req.raw);
    if (caller === null) return c.json({ error: "sign in with a valid credential" }, 401);
    c.set("who", { id: caller.id, kind: caller.kind });
    c.set("scopes", caller.scopes);
    await next();
  });
  for (const [method, path, name] of ROUTES) {
    const op = OPS.find((o) => o.name === name)!;
    api.on(method, path, async (c) => {
      const book = c.req.param("book");
      const action = op.readOnly ? "read" : "write";
      // `list` and `new` name no book, so the scope is all that limits them.
      const scopes = c.get("scopes");
      if (book === undefined && scopes !== undefined && !hasScope(scopes, `workbook:${action}`)) {
        return c.json({ error: `this credential has no ${action} scope` }, 403);
      }
      const denied = book === undefined ? null : await guard(c, action, book);
      if (denied !== null) return denied;
      let body: unknown = {};
      // A read and a delete carry their input in the path and the query string.
      if (method !== "get" && method !== "delete") {
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
    const denied = await guard(c, "admin", c.req.param("book"));
    if (denied !== null) return denied;
    const parsed = shareBody.safeParse(await c.req.json().catch(() => null));
    if (!parsed.success) return c.json({ error: message(parsed.error) }, 400);
    await access.share(c.get("who"), c.req.param("book"), { kind: parsed.data.kind, id: parsed.data.subject }, parsed.data.role as ShareRole);
    return c.json({ ok: true });
  });
  api.get("/:book/shares", async (c) => {
    const denied = await guard(c, "admin", c.req.param("book"));
    if (denied !== null) return denied;
    const query = z.object({ limit: z.coerce.number().int().min(1).max(100).optional(), cursor: z.string().max(512).optional() }).strict().safeParse(c.req.query());
    if (!query.success) return c.json({ error: message(query.error) }, 400);
    return c.json(await access.list(c.get("who"), c.req.param("book"), query.data));
  });
  api.delete("/:book/share/:kind/:subject", async (c) => {
    const kind = z.enum(SUBJECT_KIND_NAMES).safeParse(c.req.param("kind"));
    if (!kind.success) return c.json({ error: "the kind is human or agent" }, 400);
    const denied = await guard(c, "admin", c.req.param("book"));
    if (denied !== null) return denied;
    await access.unshare(c.get("who"), c.req.param("book"), { kind: kind.data, id: c.req.param("subject") }).catch((error: unknown) => {
      if (error instanceof NoShare) throw new OpError(error.message, 404);
      throw error;
    });
    return c.json({ ok: true });
  });
  // The owner deletes a book. Its grants, invites, links, and audit entries go with it.
  api.delete("/:book", async (c) => {
    const denied = await guard(c, "admin", c.req.param("book"));
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
    // The D1 adapter throws a plain Error when D1 serves a query from outside the claimed region. That is a deployment fault, not a bad request.
    if (/outside the declared region/.test(error.message)) return c.json({ error: "the database is outside the declared region" }, 500);
    if (error instanceof SheetError || error.constructor === Error) return c.json({ error: error.message }, 400);
    console.error("sheets api error", error.name);
    return c.json({ error: "something failed on the server" }, 500);
  });
  return app;
}

// The region claim. The bucket is declared with the "eu" jurisdiction in wrangler.jsonc: keep the two in step.
function regionOf(env: WorkerEnv): Region {
  let region: Region;
  try {
    region = parseRegion(env.SHEETS_REGION);
  } catch {
    throw new ConfigError('set SHEETS_REGION to "eu" or "us" ("unknown" only for local development)');
  }
  if (region === "unknown" && env.SHEETS_UNSAFE_REGION !== "1") {
    throw new ConfigError('the region "unknown" turns the residency check off; set SHEETS_UNSAFE_REGION=1 to accept that for local development');
  }
  return region;
}

// One store and one blob client for each D1 binding and region, so a transaction queue and the schema check live as long as the isolate.
const storages = new WeakMap<object, { region: Region; ready: Promise<{ store: Store; blobs: Blobs }> }>();
// The adapter needs a signing secret. This package never makes signed URLs, so a secret that lives and dies with the isolate is enough.
let urlSecret: string | undefined; // made on first use: a Worker may not make random values at load

// The store over D1 and the blobs over R2. It throws ConfigError when the bindings or the region are not set.
// The region is a claim. D1 reports where it served a query, and the store refuses a mismatch. An R2 binding
// cannot report its jurisdiction, so nothing checks the bucket against the claim: keep wrangler.jsonc in step.
export async function storageFor(env: WorkerEnv): Promise<{ store: Store; blobs: Blobs }> {
  const { DB: db, BLOBS: bucket } = env;
  if (!db || !bucket) throw new ConfigError("bind a D1 database as DB and an R2 bucket as BLOBS");
  const region = regionOf(env);
  let entry = storages.get(db);
  if (entry?.region !== region) {
    const ready = (async () => {
      await db.prepare(D1_SCHEMA).run();
      const store = createD1Store(db, { region });
      const blobs = createR2Blobs(bucket, { region, urlBase: "https://invalid.example/unused", urlSecret: (urlSecret ??= `${crypto.randomUUID()}${crypto.randomUUID()}`) });
      return { store, blobs };
    })();
    entry = { region, ready };
    storages.set(db, entry);
    ready.catch(() => storages.get(db) === entry && storages.delete(db));
  }
  return entry.ready;
}

let cached: { keys?: string; unsafe?: string; store: Store; auth: Promise<Authenticator> } | undefined;

// The authenticator for this env and store. With no key setting it reads the identity store alone. It throws ConfigError when SHEETS_API_KEYS is not valid, when both modes are set, or when the flag is not 1.
// The keys are written into the store once for each store, not on every request. The cache holds the promise,
// so requests that arrive together share one seed.
async function authenticatorFor(env: WorkerEnv, store: Store): Promise<Authenticator> {
  if (cached?.keys === env.SHEETS_API_KEYS && cached?.unsafe === env.SHEETS_UNSAFE_TRUST_HEADERS && cached?.store === store) return cached.auth;
  const entry = { keys: env.SHEETS_API_KEYS, unsafe: env.SHEETS_UNSAFE_TRUST_HEADERS, store, auth: authFromEnv(env, store) };
  cached = entry;
  try {
    return await entry.auth;
  } catch (error) {
    if (cached === entry) cached = undefined;
    throw error;
  }
}

// The role a key holds in a room. A key without the write scope joins as a viewer, and one without the admin scope never as an owner.
export function roomRole(role: Role, scopes: string[] | undefined): Role {
  const may = (action: string) => scopes === undefined || hasScope(scopes, `workbook:${action}`);
  return !may("write") ? "viewer" : role === "owner" && !may("admin") ? "editor" : role;
}

// The subprotocol of the presence room. A browser offers `agentic-spreadsheet` and `agentic-spreadsheet.bearer.<token>`,
// and the answer names the first one only.
export const WS_PROTOCOL = "agentic-spreadsheet";

// Who may join the presence room of a book, and as what. A browser cannot set a header on a WebSocket, so the key comes
// in a subprotocol entry, and the Authorization header still works for other clients. A `token` query parameter is
// not read: URLs reach logs and proxies. The room carries presence only: cells travel through the API. A principal with no read right on the book gets null.
export async function roomGrant(request: Request, book: string, env: WorkerEnv) {
  const headers = new Headers(request.headers);
  const { store } = await storageFor(env);
  const authenticate = await authenticatorFor(env, store);
  // The helper reads the subprotocol and calls this port with the key as a bearer header. The caller keeps its scopes.
  const seen = new WeakMap<PrincipalContext, Awaited<ReturnType<Authenticator>>>();
  const port: Auth = {
    async authenticate(req) {
      const found = await authenticate(req);
      if (found === null) return null;
      const context: PrincipalContext = { principalId: found.id, kind: found.kind, chain: [], scopes: found.scopes ?? [] };
      seen.set(context, found);
      return context;
    },
  };
  const context = await principalFromWebSocketRequest(new Request(request.url, { method: request.method, headers }), port, { protocol: WS_PROTOCOL });
  // A failed credential read is an error, not a wrong key: a 5xx throws, and a 401 or 403 is no access.
  if (context instanceof Response && context.status >= 500) throw new Error(`the room could not check the key (status ${context.status})`);
  const who = context instanceof Response ? null : (seen.get(context) ?? null);
  if (who === null) return null;
  const principal = { id: who.id, kind: who.kind };
  const { allowed, role } = await createAccess(store).role(principal, "read", book);
  if (!allowed || role === null || (who.scopes !== undefined && !hasScope(who.scopes, "workbook:read"))) return null;
  const capped = roomRole(role, who.scopes);
  return { principal, role: capped, name: who.id };
}

// Data lives in D1 and R2, so it outlives the isolate.
export default {
  async fetch(request: Request, env: WorkerEnv): Promise<Response> {
    let authenticate: Authenticator;
    let storage: { store: Store; blobs: Blobs };
    try {
      storage = await storageFor(env);
      authenticate = await authenticatorFor(env, storage.store);
    } catch (error) {
      if (!(error instanceof ConfigError)) throw error;
      return Response.json({ error: `the API cannot start: ${error.message}` }, { status: 500 });
    }
    return createApp({ ...storage, authenticate, pack: env.PACK_NAME }).fetch(request);
  },
};
