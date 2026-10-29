import { spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { afterEach, beforeEach, describe, expect, test } from "vitest";
import * as Y from "yjs";
import { bearerAuth } from "./auth.js";
import { FileStore } from "./core/file-store.js";
import { fromBase64, toBase64 } from "./core/bytes.js";
import { MemoryStore } from "./core/store.js";
import { createApp } from "./worker.js";
import { legacyState } from "./testing/legacy.js";

// A stored workbook in the A1-keyed form goes through the real CLI, the real MCP server over stdio,
// and the Worker app. Failure modes, listed before the code:
// - a read of an old workbook fails, or shows other values, formulas, or authors than the old version did
// - a read saves the conversion, so a plain read writes to the store
// - the first write does not save the converted form, or loses a cell or a formula on the way
// - a later insert still rewrites refs wrongly, or the CLI, MCP, and Worker give different answers
// - a stored workbook with an unaccepted old fork reads as good data, or fails with a stack trace

const CLI = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
let dir: string;
beforeEach(() => void (dir = mkdtempSync(join(tmpdir(), "sheet-legacy-"))));
afterEach(() => rmSync(dir, { recursive: true, force: true }));

const put = (store: { put(c: string, id: string, doc: Record<string, unknown>): Promise<void> }, id: string, state: Uint8Array) =>
  store.put("workbooks", id, { name: id, state: toBase64(state), updatedAt: "2026-01-01T00:00:00.000Z" });

function cli(...args: string[]) {
  const env = { ...process.env, SHEET_DIR: dir, SHEET_AS: "cli", SHEET_KIND: "agent" } as Record<string, string>;
  const r = spawnSync(process.execPath, [CLI, ...args], { encoding: "utf8", env });
  return { code: r.status, out: r.stdout, err: r.stderr, json: () => JSON.parse(r.stdout) };
}

// The sheets as they sit in the stored file.
function stored(id: string): { id?: string; cells: Record<string, unknown> }[] {
  const file = JSON.parse(readFileSync(join(dir, "workbooks", `${id}.json`), "utf8")) as { state: string };
  const doc = new Y.Doc();
  Y.applyUpdate(doc, fromBase64(file.state));
  return doc.getArray("sheets").toJSON();
}

describe("a stored old workbook", () => {
  test("the CLI reads it, a read saves nothing, and the first write saves the converted form", async () => {
    await put(new FileStore(dir), "old", legacyState());
    const read = cli("read", "old", "A1:B3").json();
    expect(read.values).toEqual([[10, 35], [20, "200A1"], ["text", null]]);
    expect(read.formulas).toEqual([[null, "=SUM(A1:A3)+'Q 3'!B2"], [null, '=$A$1*A2&"A1"'], [null, null]]);
    expect(read.authors[1][0]).toMatchObject({ id: "a1", kind: "agent_of_human", machine: true });
    expect(stored("old").every((sheet) => sheet.id === undefined)).toBe(true);
    expect(cli("export", "old").out.split("\r\n")[0]).toBe("10,35,,");
    expect(cli("insert-rows", "old", "2").code).toBe(0);
    expect(stored("old").every((sheet) => typeof sheet.id === "string")).toBe(true);
    const after = cli("read", "old", "A1:B4").json();
    expect(after.values).toEqual([[10, 35], [null, null], [20, "200A1"], ["text", null]]);
    expect(after.formulas[2][1]).toBe('=$A$1*A3&"A1"');
    expect(cli("read", "old", "C6:C6").json().formulas).toEqual([["=A3"]]);
  });

  test("the MCP server and the Worker app give the same answers as the CLI", async () => {
    await put(new FileStore(dir), "old", legacyState());
    const expected = cli("read", "old", "A1:B3").json();
    const client = new Client({ name: "test", version: "0" });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [CLI, "mcp"], env: { SHEET_DIR: dir, SHEET_AS: "bot", SHEET_KIND: "agent" } }));
    try {
      const r = (await client.callTool({ name: "sheet_read", arguments: { book: "old", range: "A1:B3" } })) as { content: { text: string }[] };
      expect(JSON.parse(r.content[0].text)).toEqual(expected);
    } finally {
      await client.close();
    }
    const key = "agent-key-0123456789-0123456789-abcdef";
    const store = new MemoryStore();
    await put(store, "old", legacyState());
    const app = createApp({ store, authenticate: await bearerAuth({ [key]: { id: "bot", kind: "agent_of_human" } }) });
    const headers = { authorization: `Bearer ${key}`, "content-type": "application/json" };
    const got = await app.request("/api/v1/sheets/old/range?range=A1:B3", { headers });
    expect(await got.json()).toEqual(expected);
    expect((await app.request("/api/v1/sheets/old/rows", { method: "POST", headers, body: JSON.stringify({ at: 2 }) })).status).toBe(200);
    const moved = await (await app.request("/api/v1/sheets/old/range?range=A1:A4", { headers })).json();
    expect(moved.values).toEqual([[10], [null], [20], ["text"]]);
  });

  test("a workbook with an unaccepted old fork is an error, not good data", async () => {
    await put(new FileStore(dir), "forked", legacyState([{ id: "a", parent: null }, { id: "b", parent: null }]));
    const r = cli("read", "forked", "A1:B3");
    expect([r.code, r.out]).toEqual([1, ""]);
    expect(r.err).toMatch(/^error: .*at the same time and not accepted/);
    expect(r.err).not.toMatch(/\n\s+at /);
  });
});
