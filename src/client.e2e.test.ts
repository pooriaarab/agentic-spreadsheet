import { spawn, type ChildProcess } from "node:child_process";
import { createServer } from "node:net";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, expectTypeOf, test } from "vitest";
import * as client from "./client.js";
import { createClient, SheetsApiError } from "./client.js";
import type { Value as EngineValue } from "./core/engine.js";
import type { Author, PrincipalKind } from "./core/workbook.js";

// These tests call the real Worker, started with `wrangler dev`, through the client.
// Failure modes, listed before the code:
// - a wrong or missing token does not reject with 401 and code unauthorized, or the message echoes the token
// - an unknown book is not a not_found error, or a key with no share on a book sees anything but not_found
// - a role that is too low is not forbidden, a second `create` is not conflict, a bad range is not bad_request
// - a method sends the wrong verb, path, or body, so a round trip of one route does not match the Worker
// - a book id with a slash or a space is not encoded, so it reaches another route
// - an aborted call resolves, or a failed call is retried (the Worker sees two requests)
// - the client file imports a Node module, so it cannot run in a browser or a Worker

const WRANGLER = fileURLToPath(new URL("../node_modules/wrangler/bin/wrangler.js", import.meta.url));
const KEY_A = "agent-key-0123456789-0123456789-abcdef";
const KEY_H = "human-key-0123456789-0123456789-abcdef";
const KEY_V = "viewer-key-0123456789-0123456789-abcde";
const KEYS = JSON.stringify({
  [KEY_A]: { id: "bot-1", kind: "agent_of_human", parent: "ana" },
  [KEY_H]: { id: "ana", kind: "human" },
  [KEY_V]: { id: "vic", kind: "human" },
});
const home = mkdtempSync(join(tmpdir(), "wrangler-client-"));

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

let stop = async () => {};
let url = "";
beforeAll(async () => {
  for (let attempt = 1; ; attempt += 1) {
    const [port, inspector] = await freePorts(2);
    const args = [WRANGLER, "dev", "--port", String(port), "--inspector-port", String(inspector), "--ip", "127.0.0.1", "--persist-to", join(home, "state"), "--var", "SHEETS_REGION:unknown", "--var", "SHEETS_UNSAFE_REGION:1", "--var", `SHEETS_API_KEYS:${KEYS}`];
    const env = { ...process.env, CI: "1", NO_COLOR: "1", WRANGLER_SEND_METRICS: "false", XDG_CONFIG_HOME: home, HOME: home };
    const child: ChildProcess = spawn(process.execPath, args, { env, stdio: ["ignore", "pipe", "pipe"], detached: true });
    let [exited, log] = [false, ""];
    child.on("exit", () => (exited = true));
    for (const stream of [child.stdout, child.stderr]) stream?.on("data", (chunk: Buffer) => (log = (log + chunk).slice(-1500)));
    const kill = async () => void (child.pid && !exited && process.kill(-child.pid, "SIGKILL"));
    const base = `http://127.0.0.1:${port}`;
    for (let i = 0; i < 300 && !exited; i += 1) {
      if (await fetch(`${base}/health`).then(() => true, () => false)) {
        [url, stop] = [base, kill];
        return;
      }
      await new Promise((r) => setTimeout(r, 300));
    }
    await kill();
    if (!exited || attempt >= 3) throw new Error(`wrangler dev did not start after ${attempt} tries: ${log}`);
  }
}, 120_000);
afterAll(async () => {
  await stop();
  rmSync(home, { recursive: true, force: true });
});

const as = (token: string) => createClient({ baseUrl: url, token });
const failure = async (call: Promise<unknown>) => {
  const error = await call.then(() => null, (e: unknown) => e);
  expect(error).toBeInstanceOf(SheetsApiError);
  return error as SheetsApiError;
};

// The client copies these types so it has no imports. Typecheck fails when a copy drifts.
test("the client types match the core types", () => {
  expectTypeOf<client.Value>().toEqualTypeOf<EngineValue>();
  expectTypeOf<client.PrincipalKind>().toEqualTypeOf<PrincipalKind>();
  expectTypeOf<client.CellAuthor>().toEqualTypeOf<Author>();
});

describe("client against the Worker", () => {
  test("health", async () => {
    expect(await as("x").health()).toMatchObject({ ok: true, pack: "agentic-spreadsheet" });
  });

  test("a round trip over every route", async () => {
    const bot = as(KEY_A);
    expect(await bot.create({ book: "trip" })).toEqual({ book: "trip", sheets: ["Sheet1"] });
    expect(await bot.share({ book: "trip", kind: "human", subject: "ana", role: "editor" })).toEqual({ ok: true });
    expect(await bot.write({ book: "trip", anchor: "A1", values: [[2, "=A1*3", "x"]] })).toEqual({ ok: true, rows: 1 });
    expect(await bot.setFormula({ book: "trip", cell: "A2", formula: "A1+1" })).toEqual({ cell: "A2", formula: "=A1+1", value: 3 });
    expect(await bot.insertRows({ book: "trip", at: 1, count: 2 })).toEqual({ ok: true });
    const read = await as(KEY_H).read({ book: "trip", range: "A3:C4" });
    expect(read.values).toEqual([[2, 6, "x"], [3, null, null]]);
    expect(read.formulas[1][0]).toBe("=A3+1");
    expect(read.authors[0][0]).toMatchObject({ id: "bot-1", kind: "agent_of_human", machine: true });
    expect(await bot.exportCsv({ book: "trip" })).toBe(",,\r\n,,\r\n2,6,x\r\n3,,\r\n");
    expect(await bot.importCsv({ book: "trip", csv: "a,b\n1,=A2+1" })).toEqual({ ok: true, rows: 2 });
    expect((await bot.read({ book: "trip", range: "A1:B2" })).values).toEqual([["a", "b"], [1, 2]]);
    expect(await bot.list()).toEqual({ books: ["trip"] });
    expect(await bot.list({ limit: 1 })).toMatchObject({ books: ["trip"] });
    const invite = await bot.invite({ book: "trip", role: "viewer", handle: "helper" });
    const link = await bot.link({ book: "trip", role: "viewer", expiresInMs: 60_000 });
    expect((await bot.invites({ book: "trip" })).invites.map((i) => i.id)).toEqual([invite.id]);
    expect((await bot.links({ book: "trip" })).links).toMatchObject([{ id: link.id, role: "viewer", signedInOnly: false }]);
    expect(await bot.unlink({ book: "trip", link: link.id })).toEqual({ ok: true });
    expect((await failure(bot.unlink({ book: "trip", link: link.id + "x" }))).code).toBe("not_found");
    expect((await failure(as(KEY_H).links({ book: "trip" }))).code).toBe("forbidden");
    expect(await bot.unshare({ book: "trip", kind: "human", subject: "ana" })).toEqual({ ok: true });
    expect(await bot.remove({ book: "trip" })).toEqual({ ok: true });
    expect((await failure(bot.read({ book: "trip", range: "A1" }))).code).toBe("not_found");
  });

  test("structure calls change rows, columns, and sheets, and a formula follows or reads #REF!", async () => {
    const bot = as(KEY_A);
    await bot.create({ book: "shape" });
    await bot.write({ book: "shape", anchor: "A1", values: [[1, 2, "=A1+B1"], [3, 4, "=A2+B2"]] });
    expect(await bot.insertCols({ book: "shape", at: 1 })).toEqual({ ok: true });
    expect((await bot.read({ book: "shape", range: "B1:D1" })).formulas[0]).toEqual([null, null, "=B1+C1"]);
    expect(await bot.deleteCols({ book: "shape", at: 1 })).toEqual({ ok: true });
    expect(await bot.deleteRows({ book: "shape", at: 1, count: 1 })).toEqual({ ok: true });
    expect((await bot.read({ book: "shape", range: "A1:C1" })).values).toEqual([[3, 4, 7]]);
    await bot.write({ book: "shape", anchor: "A3", values: [["=A1"]] });
    await bot.deleteRows({ book: "shape", at: 1 });
    expect((await bot.read({ book: "shape", range: "A2:A2" })).values).toEqual([[{ error: "#REF!" }]]);
    expect(await bot.addSheet({ book: "shape", name: "Data 2" })).toEqual({ ok: true, sheets: ["Sheet1", "Data 2"] });
    await bot.write({ book: "shape", anchor: "A1", values: [[5]], sheet: "Data 2" });
    expect(await bot.sheets({ book: "shape" })).toEqual({ sheets: [{ name: "Sheet1", extent: "A1:A2" }, { name: "Data 2", extent: "A1:A1" }] });
    expect(await bot.deleteSheet({ book: "shape", name: "Data 2" })).toEqual({ ok: true, sheets: ["Sheet1"] });
    expect(await failure(bot.deleteSheet({ book: "shape", name: "Sheet1" }))).toMatchObject({ status: 400, code: "bad_request" });
    expect(await failure(bot.addSheet({ book: "shape", name: "Sheet1" }))).toMatchObject({ status: 400 });
    expect(await failure(bot.deleteRows({ book: "shape", at: 0 }))).toMatchObject({ status: 400 });
    // A delete past the end of the sheet is an error, and so is a sheet name that is too long.
    expect(await failure(bot.deleteRows({ book: "shape", at: 1_048_576, count: 2 }))).toMatchObject({ status: 400 });
    expect(await failure(bot.deleteCols({ book: "shape", at: 16_384, count: 2 }))).toMatchObject({ status: 400 });
    expect(await failure(bot.addSheet({ book: "shape", name: "x".repeat(257) }))).toMatchObject({ status: 400 });
  });

  test("a viewer cannot change the structure, and a stranger sees no workbook", async () => {
    const bot = as(KEY_A);
    await bot.create({ book: "shape2" });
    await bot.share({ book: "shape2", kind: "human", subject: "vic", role: "viewer" });
    const viewer = as(KEY_V);
    expect((await viewer.sheets({ book: "shape2" })).sheets).toHaveLength(1);
    const attempts = [
      () => viewer.addSheet({ book: "shape2", name: "X" }),
      () => viewer.deleteRows({ book: "shape2", at: 1 }),
      () => viewer.insertCols({ book: "shape2", at: 1 }),
      () => viewer.deleteSheet({ book: "shape2", name: "Sheet1" }),
    ];
    for (const attempt of attempts) {
      expect(await failure(attempt())).toMatchObject({ status: 403, code: "forbidden" });
    }
    expect((await failure(as(KEY_H).sheets({ book: "shape2" }))).code).toBe("not_found");
    expect((await failure(as(KEY_H).deleteCols({ book: "shape2", at: 1 }))).code).toBe("not_found");
  });

  test("a failed formula reads as an error value", async () => {
    const bot = as(KEY_A);
    await bot.create({ book: "errs" });
    expect(await bot.setFormula({ book: "errs", cell: "A1", formula: "1/0" })).toMatchObject({ value: { error: "#DIV/0!" } });
    await bot.write({ book: "errs", anchor: "B1", values: [["=A1+1"]] });
    await bot.insertRows({ book: "errs", at: 1 });
    expect((await bot.read({ book: "errs", range: "A2:B2" })).values).toEqual([[{ error: "#DIV/0!" }, { error: "#DIV/0!" }]]);
  });

  test("health keeps the message of a failed start", async () => {
    expect(await failure(createClient({ baseUrl: url, token: "x", fetch: async () => Response.json({ error: "the API cannot start: x" }, { status: 500 }) }).health()))
      .toMatchObject({ status: 500, code: "server_error", message: "the API cannot start: x" });
  });

  test("a wrong token is unauthorized, and the message does not echo it", async () => {
    const token = "x".repeat(KEY_A.length);
    const error = await failure(as(token).list());
    expect([error.status, error.code]).toEqual([401, "unauthorized"]);
    expect(error.message).not.toContain(token);
  });

  test("an unknown book is not_found, and so is a book with no share", async () => {
    const missing = await failure(as(KEY_A).read({ book: "nope", range: "A1" }));
    expect([missing.status, missing.code, missing.message]).toEqual([404, "not_found", 'no workbook "nope"']);
    await as(KEY_A).create({ book: "private" });
    expect((await failure(as(KEY_H).read({ book: "private", range: "A1" }))).code).toBe("not_found");
  });

  test("forbidden, conflict, and bad_request carry their code", async () => {
    const bot = as(KEY_A);
    await bot.create({ book: "roles" });
    await bot.share({ book: "roles", kind: "human", subject: "vic", role: "viewer" });
    expect(await failure(as(KEY_V).write({ book: "roles", anchor: "A1", values: [[1]] }))).toMatchObject({ status: 403, code: "forbidden" });
    expect(await failure(bot.create({ book: "roles" }))).toMatchObject({ status: 409, code: "conflict" });
    expect(await failure(bot.read({ book: "roles", range: "not a range" }))).toMatchObject({ status: 400, code: "bad_request" });
  });

  test("a book id with a slash is one path segment, and the Worker names it whole", async () => {
    expect((await failure(as(KEY_A).read({ book: "a/b", range: "A1" }))).message).toBe('no workbook "a/b"');
    expect(await failure(as(KEY_A).create({ book: "a b" }))).toMatchObject({ status: 400, code: "bad_request" });
  });

  test("an aborted call rejects and the Worker is not called twice", async () => {
    const seen: string[] = [];
    const counting = createClient({ baseUrl: url, token: KEY_A, fetch: (input, init) => (seen.push(String(input)), fetch(input, init)) });
    const controller = new AbortController();
    controller.abort();
    await expect(counting.list({}, { signal: controller.signal })).rejects.toMatchObject({ name: "AbortError" });
    // A failed call is not retried either.
    await failure(counting.read({ book: "nope", range: "A1" }));
    expect(seen.filter((u) => u.includes("/nope/"))).toHaveLength(1);
  });

  test("the client file has no Node import", () => {
    const source = readFileSync(fileURLToPath(new URL("./client.ts", import.meta.url)), "utf8");
    expect(source).not.toMatch(/^\s*import\s|require\(|\bBuffer\b|\bprocess\./m);
  });
});
