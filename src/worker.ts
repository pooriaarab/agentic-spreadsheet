import { Hono } from "hono";
import { bodyLimit } from "hono/body-limit";
import { z } from "zod";
import { authFromEnv, ConfigError, type Authenticator, type AuthEnv } from "./auth.js";
import { SheetError, type Principal } from "./core/workbook.js";
import { MemoryStore, type Store } from "./core/store.js";
import { message, OpError, OPS } from "./ops.js";
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
  authenticate: Authenticator;
  pack?: string;
}

export function createApp({ store, authenticate, pack = "agentic-spreadsheet" }: AppOptions): Hono {
  const app = new Hono();
  app.use(async (c, next) => {
    await next();
    c.header("cache-control", "no-store");
    c.header("x-content-type-options", "nosniff");
  });
  app.get("/health", (c) => c.json({ ok: true, pack, version: VERSION }));

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
      let body: unknown = {};
      if (method !== "get") {
        body = await c.req.json().catch(() => null);
        if (typeof body !== "object" || body === null || Array.isArray(body)) return c.json({ error: "the body must be a JSON object" }, 400);
      }
      const parsed = z.object(op.shape).strict().safeParse({ ...c.req.query(), ...(body as object), ...c.req.param() });
      if (!parsed.success) return c.json({ error: message(parsed.error) }, 400);
      const result = await op.run(parsed.data, { store, principal: () => c.get("who"), maxWriteCells: MAX_WRITE_CELLS });
      if (name === "export") return c.body((result as { csv: string }).csv, 200, { "content-type": "text/csv; charset=utf-8" });
      return c.json(result as object, name === "new" ? 201 : 200);
    });
  }
  app.route("/api/v1/sheets", api);

  app.notFound((c) => c.json({ error: "not found" }, 404));
  // A message from an operation is for the caller. Anything else is a bug: log its name, not its text.
  app.onError((error, c) => {
    if (error instanceof OpError) return c.json({ error: error.message }, error.status);
    if (error instanceof SheetError || error.constructor === Error) return c.json({ error: error.message }, 400);
    console.error("sheets api error", error.name);
    return c.json({ error: "something failed on the server" }, 500);
  });
  return app;
}

const store = new MemoryStore();
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
    return createApp({ store, authenticate, pack: env.PACK_NAME }).fetch(request);
  },
};
