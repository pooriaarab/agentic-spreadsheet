import { spawn, type ChildProcess } from "node:child_process";
import { createServer } from "node:net";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, test } from "vitest";

// These tests start the real Worker with `wrangler dev` and call it over HTTP.
// Failure modes, listed before the code:
// - the Worker serves with no auth set, with both modes set, with a key under 32 bytes, with a key
//   that has no kind or a bad kind, or with bad JSON in SHEETS_API_KEYS
// - a missing, wrong, or near-miss token gets in, or an error echoes the token
// - in header mode a missing or bad kind gets in, or an agent edit is not labeled machine-made
// - a body field can change who the edit belongs to
// - twenty parallel writes to one book lose one
// - an unknown book is not 404, a second `new` is not 409, bad JSON or an unknown field is not 400,
//   a 6 MB body is not 413, a count of 2,000,000 is not 400
// - one deep formula turns a read into a 500
// - one 5 MB write of about 2 million cells, or a 10,001-cell write or import, is accepted
// - a formula over ten 1,000,000-cell ranges is computed in full
// - two far-apart cells (A1 and XFD1048576) make export build a huge grid, or kill the Worker
// - export is not CSV, or a response can be cached
// - /health reports a version other than package.json

const WRANGLER = fileURLToPath(new URL("../node_modules/wrangler/bin/wrangler.js", import.meta.url));
const VERSION = (JSON.parse(readFileSync(fileURLToPath(new URL("../package.json", import.meta.url)), "utf8")) as { version: string }).version;
const KEY_A = "agent-key-0123456789-0123456789-abcdef";
const KEY_H = "human-key-0123456789-0123456789-abcdef";
const KEYS = JSON.stringify({ [KEY_A]: { id: "bot-1", kind: "agent_of_human" }, [KEY_H]: { id: "ana", kind: "human" } });

interface Server {
  url: string;
  stop(): Promise<void>;
}
const home = mkdtempSync(join(tmpdir(), "wrangler-"));

// Ports that are free now. They are all open together, so no two are the same.
async function freePorts(count: number): Promise<number[]> {
  const open = await Promise.all(
    Array.from({ length: count }, () => new Promise<ReturnType<typeof createServer>>((resolve) => {
      const server = createServer().listen(0, "127.0.0.1", () => resolve(server));
    })),
  );
  const ports = open.map((server) => (server.address() as { port: number }).port);
  await Promise.all(open.map((server) => new Promise((resolve) => server.close(resolve))));
  return ports;
}

// Each server gets its own port and its own inspector port. The inspector port defaults to one
// number for every server, and two servers that start together clash on it. A port that was free
// can still be taken before wrangler binds it, so a server that exits at start is started again, up
// to three times.
async function start(vars: Record<string, string>): Promise<Server> {
  for (let attempt = 1; ; attempt += 1) {
    const [port, inspector] = await freePorts(2);
    const args = [WRANGLER, "dev", "--port", String(port), "--inspector-port", String(inspector), "--ip", "127.0.0.1", ...Object.entries(vars).flatMap(([k, v]) => ["--var", `${k}:${v}`])];
    const env = { ...process.env, CI: "1", NO_COLOR: "1", WRANGLER_SEND_METRICS: "false", XDG_CONFIG_HOME: home, HOME: home };
    const child: ChildProcess = spawn(process.execPath, args, { env, stdio: ["ignore", "pipe", "pipe"], detached: true });
    let [exited, log] = [false, ""];
    child.on("exit", () => (exited = true));
    for (const stream of [child.stdout, child.stderr]) stream?.on("data", (chunk: Buffer) => (log = (log + chunk).slice(-1500)));
    const url = `http://127.0.0.1:${port}`;
    const stop = async () => void (child.pid && !exited && process.kill(-child.pid, "SIGKILL"));
    for (let i = 0; i < 300; i += 1) {
      if (exited) break;
      if (await fetch(`${url}/health`).then(() => true, () => false)) return { url, stop };
      await new Promise((r) => setTimeout(r, 300));
    }
    await stop();
    if (!exited || attempt >= 3) throw new Error(`wrangler dev did not start after ${attempt} tries: ${log}`);
  }
}

const servers: Record<string, Server> = {};
beforeAll(async () => {
  const configs: Record<string, Record<string, string>> = {
    none: {},
    both: { SHEETS_API_KEYS: KEYS, SHEETS_UNSAFE_TRUST_HEADERS: "1" },
    short: { SHEETS_API_KEYS: JSON.stringify({ "too-short": { id: "a", kind: "agent" } }) },
    nokind: { SHEETS_API_KEYS: JSON.stringify({ [KEY_A]: { id: "a" } }) },
    badkind: { SHEETS_API_KEYS: JSON.stringify({ [KEY_A]: { id: "a", kind: "root" } }) },
    badjson: { SHEETS_API_KEYS: "{not json" },
    keys: { SHEETS_API_KEYS: KEYS },
    headers: { SHEETS_UNSAFE_TRUST_HEADERS: "1" },
  };
  await Promise.all(Object.entries(configs).map(async ([name, vars]) => void (servers[name] = await start(vars))));
}, 180_000);
afterAll(async () => {
  await Promise.all(Object.values(servers).map((s) => s.stop()));
  rmSync(home, { recursive: true, force: true });
});

const auth = (key: string) => ({ authorization: `Bearer ${key}` });
async function call(server: string, method: string, path: string, init: { headers?: Record<string, string>; body?: unknown; raw?: string } = {}) {
  const res = await fetch(`${servers[server].url}${path}`, {
    method,
    headers: { ...(init.body !== undefined || init.raw !== undefined ? { "content-type": "application/json" } : {}), ...init.headers },
    body: init.raw ?? (init.body === undefined ? undefined : JSON.stringify(init.body)),
  });
  const text = await res.text();
  return { status: res.status, headers: res.headers, text, json: () => JSON.parse(text) };
}
const API = "/api/v1/sheets";
const as = (key: string) => ({ headers: auth(key) });

describe("refuses to start without safe auth", () => {
  test.each([
    ["none", /no auth is set.*SHEETS_API_KEYS.*SHEETS_UNSAFE_TRUST_HEADERS/],
    ["both", /not both/],
    ["short", /at least 32 bytes/],
    ["nokind", /id and a kind/],
    ["badkind", /id and a kind/],
    ["badjson", /must be JSON/],
  ])("%s", async (name, message) => {
    for (const path of ["/health", API, `${API}/b/range?range=A1`]) {
      const res = await call(name, "GET", path, { headers: { "x-principal-id": "x", "x-principal-kind": "human", ...auth(KEY_A) } });
      expect(res.status).toBe(500);
      expect(res.json().error).toMatch(message);
      expect(res.text).not.toContain(KEY_A);
    }
  });
});

describe("bearer keys", () => {
  test.each([
    ["no header", {}], ["wrong key", auth("x".repeat(KEY_A.length))], ["near miss", auth(`${KEY_A}x`)], ["short", auth("a")],
    ["wrong scheme", { authorization: `Basic ${KEY_A}` }], ["header mode headers", { "x-principal-id": "a", "x-principal-kind": "human" }],
  ])("%s gets 401 and the response does not echo the token", async (_name, headers) => {
    const res = await call("keys", "GET", API, { headers });
    expect([res.status, res.text.includes(KEY_A)]).toEqual([401, false]);
    expect(res.headers.get("cache-control")).toBe("no-store");
  });

  test("a valid key works, and the edit carries the key's principal and kind", async () => {
    expect((await call("keys", "POST", API, { ...as(KEY_A), body: { book: "b" } })).status).toBe(201);
    expect((await call("keys", "PUT", `${API}/b/share`, { ...as(KEY_A), body: { kind: "human", subject: "ana", role: "editor" } })).status).toBe(200);
    expect((await call("keys", "PUT", `${API}/b/range`, { ...as(KEY_A), body: { anchor: "A1", values: [[2, "=A1*3"]] } })).json()).toEqual({ ok: true, rows: 1 });
    await call("keys", "PUT", `${API}/b/range`, { ...as(KEY_H), body: { anchor: "A2", values: [[5]] } });
    const read = (await call("keys", "GET", `${API}/b/range?range=A1:B2`, as(KEY_H))).json();
    expect(read.values).toEqual([[2, 6], [5, null]]);
    expect(read.authors[0][0]).toMatchObject({ id: "bot-1", kind: "agent_of_human", machine: true });
    expect(read.authors[1][0]).toMatchObject({ id: "ana", kind: "human", machine: false });
    expect((await call("keys", "GET", API, as(KEY_H))).json()).toEqual({ books: ["b"] });
  });

  test("a body field cannot change who the edit belongs to", async () => {
    await call("keys", "POST", API, { ...as(KEY_A), body: { book: "own" } });
    for (const extra of [{ kind: "human" }, { principal: { id: "x", kind: "human" } }, { as: "ana" }]) {
      const res = await call("keys", "PUT", `${API}/own/range`, { ...as(KEY_A), body: { anchor: "A1", values: [[1]], ...extra } });
      expect(res.status).toBe(400);
    }
  });

  test("twenty parallel writes to one book all land", async () => {
    await call("keys", "POST", API, { ...as(KEY_A), body: { book: "par" } });
    const codes = await Promise.all(
      Array.from({ length: 20 }, (_, i) => call("keys", "PUT", `${API}/par/range`, { ...as(KEY_A), body: { anchor: `A${i + 1}`, values: [[i + 1]] } }).then((r) => r.status)),
    );
    expect(codes).toEqual(Array(20).fill(200));
    const read = (await call("keys", "GET", `${API}/par/range?range=A1:A20`, as(KEY_A))).json();
    expect(read.values.flat()).toEqual(Array.from({ length: 20 }, (_, i) => i + 1));
  });

  test("errors have the right status and no stack", async () => {
    await call("keys", "POST", API, { ...as(KEY_A), body: { book: "e" } });
    const cases: [string, string, Parameters<typeof call>[3], number][] = [
      ["GET", `${API}/nope/range?range=A1`, as(KEY_A), 404],
      ["POST", API, { ...as(KEY_A), body: { book: "e" } }, 409],
      ["POST", API, { ...as(KEY_A), body: { book: "../x" } }, 400],
      ["POST", API, { ...as(KEY_A), raw: "{not json" }, 400],
      ["POST", API, { ...as(KEY_A), raw: "[1]" }, 400],
      ["PUT", `${API}/e/range`, { ...as(KEY_A), body: { anchor: "A1", values: "x", extra: 1 } }, 400],
      ["GET", `${API}/e/range?range=A0`, as(KEY_A), 400],
      ["POST", `${API}/e/rows`, { ...as(KEY_A), body: { at: 1, count: 2_000_000 } }, 400],
      ["PUT", `${API}/e/formula`, { ...as(KEY_A), body: { cell: "A1", formula: `=${"1+".repeat(5000)}1` } }, 400],
      ["PUT", `${API}/e/range`, { ...as(KEY_A), raw: JSON.stringify({ anchor: "A1", values: [["x".repeat(6 * 1024 * 1024)]] }) }, 413],
      ["DELETE", `${API}/e/range`, as(KEY_A), 404],
      ["GET", "/api/elsewhere", {}, 404],
    ];
    for (const [method, path, init, status] of cases) {
      const res = await call("keys", method, path, init);
      expect([method, path.slice(0, 40), res.status]).toEqual([method, path.slice(0, 40), status]);
      expect(res.text).not.toMatch(/\n\s+at |node_modules|TypeError/);
    }
  });

  test("a deep formula is one cell error, and export is CSV", async () => {
    await call("keys", "POST", API, { ...as(KEY_A), body: { book: "d" } });
    await call("keys", "PUT", `${API}/d/range`, { ...as(KEY_A), body: { anchor: "A1", values: [[1, `=${"(".repeat(4000)}1${")".repeat(4000)}`, "a,b", "+cmd"]] } });
    expect((await call("keys", "GET", `${API}/d/range?range=A1:B1`, as(KEY_A))).json().values).toEqual([[1, { error: "#DEPTH!" }]]);
    const csv = await call("keys", "GET", `${API}/d/export`, as(KEY_A));
    expect([csv.status, csv.headers.get("content-type"), csv.text]).toEqual([200, "text/csv; charset=utf-8", `1,#DEPTH!,"a,b",'+cmd\r\n`]);
    expect((await call("keys", "POST", `${API}/d/import`, { ...as(KEY_A), body: { csv: "=1+1,x\n" } })).status).toBe(200);
    expect((await call("keys", "GET", `${API}/d/range?range=A1:B1`, as(KEY_A))).json().values).toEqual([[2, "x"]]);
  });

  test("a write or import over 10,000 cells is 413, and the Worker keeps serving", { timeout: 60_000 }, async () => {
    await call("keys", "POST", API, { ...as(KEY_A), body: { book: "cap" } });
    const put = (cells: number) => call("keys", "PUT", `${API}/cap/range`, { ...as(KEY_A), body: { anchor: "A1", values: [Array(cells).fill(1)] } });
    expect((await put(10_000)).status).toBe(200);
    const over = await put(10_001);
    expect([over.status, over.json().error]).toEqual([413, expect.stringMatching(/10,000 cells/)]);
    // About 2 million cells in a 4 MB body. It used to need more memory than the Worker has.
    const big = await call("keys", "PUT", `${API}/cap/range`, { ...as(KEY_A), raw: `{"anchor":"A1","values":[${Array(122).fill(`[${Array(16_384).fill(1).join(",")}]`).join(",")}]}` });
    expect(big.status).toBe(413);
    const csv = Array.from({ length: 101 }, () => Array(100).fill(1).join(",")).join("\n");
    expect((await call("keys", "POST", `${API}/cap/import`, { ...as(KEY_A), body: { csv } })).status).toBe(413);
    expect((await call("keys", "GET", "/health")).status).toBe(200);
    expect((await call("keys", "GET", `${API}/cap/range?range=A1:B1`, as(KEY_A))).json().values).toEqual([[1, 1]]);
  });

  test("a formula over ten 1,000,000-cell ranges is one #VALUE! cell, and export of a sparse wide sheet is 413", { timeout: 60_000 }, async () => {
    await call("keys", "POST", API, { ...as(KEY_A), body: { book: "work" } });
    const ranges = "ABCDEFGHIJ".split("").map((c) => `${c}1:${c}1000000`).join(",");
    await call("keys", "PUT", `${API}/work/formula`, { ...as(KEY_A), body: { cell: "L1", formula: `=SUM(${ranges})` } });
    expect((await call("keys", "GET", `${API}/work/range?range=L1`, as(KEY_A))).json().values).toEqual([[{ error: "#VALUE!" }]]);
    await call("keys", "PUT", `${API}/work/range`, { ...as(KEY_A), body: { anchor: "XFD1048576", values: [[2]] } });
    const started = Date.now();
    const res = await call("keys", "GET", `${API}/work/export`, as(KEY_A));
    expect([res.status, res.json().error]).toEqual([413, expect.stringMatching(/too large to export/)]);
    expect(Date.now() - started).toBeLessThan(10_000);
    expect((await call("keys", "GET", "/health")).status).toBe(200);
  });

  test("/health needs no key and reports the package version", async () => {
    const res = await call("keys", "GET", "/health");
    expect([res.status, res.json()]).toEqual([200, { ok: true, pack: "agentic-spreadsheet", version: VERSION }]);
  });
});

describe("header mode (SHEETS_UNSAFE_TRUST_HEADERS=1)", () => {
  test.each([
    ["no headers", {}], ["no kind", { "x-principal-id": "a" }], ["no id", { "x-principal-kind": "agent" }],
    ["bad kind", { "x-principal-id": "a", "x-principal-kind": "root" }], ["bad id", { "x-principal-id": "a b", "x-principal-kind": "agent" }],
  ])("%s gets 401", async (_name, headers) => {
    expect((await call("headers", "GET", API, { headers })).status).toBe(401);
  });

  test("a principal from the headers labels the edit", async () => {
    const who = { "x-principal-id": "tool-7", "x-principal-kind": "agent" };
    await call("headers", "POST", API, { headers: who, body: { book: "h" } });
    await call("headers", "PUT", `${API}/h/range`, { headers: who, body: { anchor: "A1", values: [[1]] } });
    const read = (await call("headers", "GET", `${API}/h/range?range=A1`, { headers: who })).json();
    expect(read.authors[0][0]).toMatchObject({ id: "tool-7", kind: "agent", machine: true });
    expect((await call("headers", "GET", API, as(KEY_A))).status).toBe(401);
  });
});
