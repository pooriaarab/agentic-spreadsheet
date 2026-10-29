import { spawn, spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { afterEach, beforeEach, describe, expect, test } from "vitest";

// These tests drive the built CLI and the built MCP server over stdio.
// Failure modes, listed before the code:
// - an id like ../x, a/b, or "" reaches the file system; a second `new` overwrites the first
// - a missing book, a bad sheet, bad JSON, or a bad kind prints a stack trace or exits 0
// - "007" loses its zeros, a formula is not computed on read, an error value prints as text
// - the CLI writes as an agent, an agent edit shows as human-made, or a bad kind is accepted
// - state is not saved between commands, or insert-rows does not rewrite refs after a reload
// - CSV export does not quote commas, quotes, or newlines; text that starts with = + - @
//   stays live in a spreadsheet app; import and export lose "007" or booleans; list ignores paging
// - the MCP server crashes on bad input, or returns other shapes than the CLI
// - twelve CLI processes, or twelve parallel MCP calls, write one book and an update is lost
// - a write without --kind or --as lands as human-made, or a bad kind is accepted
// - a sheet of 150,000 rows cannot be imported or exported
// - insert-rows takes a count of 2,000,000, or a formula over 8,192 characters is saved
// - one deep formula makes read or export fail for the whole range
// - text that starts with + - @ gains an apostrophe on every round trip; a BOM stays in A1;
//   a 20-digit id loses digits; --as-text still runs a leading = as a formula
// - two far-apart cells (A1 and XFD1048576) make export build an 82-million-cell grid
// - list restarts at page 1 on a bad cursor

const CLI = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
let dir: string;
beforeEach(() => void (dir = mkdtempSync(join(tmpdir(), "sheet-"))));
afterEach(() => rmSync(dir, { recursive: true, force: true }));

// The principal always comes from the caller. Tests that write say who they are.
const who = (): Record<string, string | undefined> => ({ SHEET_DIR: dir, SHEET_AS: "cli", SHEET_KIND: "human" });
function run(env: Record<string, string | undefined>, args: string[]) {
  const clean = Object.fromEntries(Object.entries({ ...process.env, ...env }).filter(([, v]) => v !== undefined)) as Record<string, string>;
  const r = spawnSync(process.execPath, [CLI, ...args], { encoding: "utf8", env: clean, maxBuffer: 64 * 1024 * 1024 });
  return { code: r.status, out: r.stdout, err: r.stderr, json: () => JSON.parse(r.stdout) };
}
const cli = (...args: string[]) => run(who(), args);
const cliAsync = (...args: string[]) =>
  new Promise<number | null>((resolve) =>
    spawn(process.execPath, [CLI, ...args], { env: { ...process.env, ...who() } as Record<string, string>, stdio: "ignore" }).on("close", resolve),
  );

describe("cli", () => {
  test("values, formulas, errors, and authors survive separate commands", () => {
    expect(cli("new", "b").code).toBe(0);
    cli("write", "b", "A1", '[[1,"007",true,"=A1+1","=1/0"]]');
    cli("write", "b", "A2", "[[5]]", "--as", "bot", "--kind", "agent_of_human");
    const read = cli("read", "b", "A1:E2").json();
    expect(read.values).toEqual([[1, "007", true, 2, { error: "#DIV/0!" }], [5, null, null, null, null]]);
    expect(read.formulas[0]).toEqual([null, null, null, "=A1+1", "=1/0"]);
    expect(read.authors[0][0]).toMatchObject({ id: "cli", kind: "human", machine: false });
    expect(read.authors[1][0]).toMatchObject({ id: "bot", kind: "agent_of_human", machine: true });
    expect(cli("formula", "b", "F1", "SUM(A1:A2)").json()).toMatchObject({ formula: "=SUM(A1:A2)", value: 6 });
  });

  test("insert-rows rewrites refs and the change persists", () => {
    cli("new", "b");
    cli("write", "b", "A1", '[[1],[2],["=SUM(A1:A2)"]]');
    expect(cli("insert-rows", "b", "2").code).toBe(0);
    const read = cli("read", "b", "A1:A4").json();
    expect(read.formulas[3][0]).toBe("=SUM(A1:A3)");
    expect(read.values).toEqual([[1], [null], [2], [3]]);
  });

  test("csv export quotes, neutralises formulas, and round-trips", () => {
    cli("new", "b");
    cli("write", "b", "A1", JSON.stringify([["a,b", 'say "hi"', "two\nlines", "+cmd", 1], ["007", true, "=A1", null, -2]]));
    const exported = cli("export", "b").out;
    expect(exported).toBe('"a,b","say ""hi""","two\nlines",\'+cmd,1\r\n007,TRUE,"a,b",,-2\r\n');
    const file = join(dir, "in.csv");
    writeFileSync(file, exported);
    cli("new", "c");
    expect(cli("import", "c", "--file", file).code).toBe(0);
    expect(cli("read", "c", "A1:E2").json().values).toEqual([
      ["a,b", 'say "hi"', "two\nlines", "+cmd", 1], ["007", true, "a,b", null, -2],
    ]);
  });

  test("list pages through books in name order", () => {
    for (const name of ["c", "a", "b"]) cli("new", name);
    const first = cli("list", "--limit", "2").json();
    expect(first).toEqual({ books: ["a", "b"], cursor: expect.any(String) });
    expect(cli("list", "--limit", "2", "--cursor", first.cursor).json()).toEqual({ books: ["c"] });
  });

  test.each([
    [["new", "../x"], 1], [["new", "a/b"], 1], [["new", ""], 1], [["read", "nope", "A1"], 1], [["new", "b"], 1],
    [["write", "b", "A1", "not json"], 1], [["write", "b", "A1", "[[1]]", "--kind", "root"], 1],
    [["read", "b", "A1", "--sheet", "Nope"], 1], [["read", "b", "A0"], 1], [["frobnicate"], 2],
    [["insert-rows", "b", "1", "2000000"], 1], [["insert-rows", "b", "2000000"], 1], [["list", "--cursor", "../x"], 1],
    [["formula", "b", "A1", `=${"1+".repeat(5000)}1`], 1],
  ])("%j fails cleanly with exit %i", (args, code) => {
    cli("new", "b");
    const r = cli(...args);
    expect(r.code).toBe(code);
    expect(r.out).toBe("");
    expect(r.err).toMatch(/^error: |unknown command/);
    expect(r.err).not.toMatch(/\n\s+at /);
  });
});

describe("who is writing", () => {
  test.each([
    [{ SHEET_KIND: undefined }, /--kind/], [{ SHEET_AS: undefined }, /--as/], [{ SHEET_KIND: "root" }, /--kind/], [{ SHEET_KIND: "" }, /--kind/],
  ])("a write with %j is refused, and a read still works", (env, message) => {
    cli("new", "b");
    const w = run({ ...who(), ...env }, ["write", "b", "A1", "[[1]]"]);
    expect([w.code, w.err]).toEqual([1, expect.stringMatching(message)]);
    expect(run({ ...who(), ...env }, ["read", "b", "A1"]).code).toBe(0);
    expect(cli("read", "b", "A1").json().values).toEqual([[null]]);
  });
});

describe("shared files", () => {
  test("twelve CLI processes write one book and every write lands", async () => {
    cli("new", "b");
    const codes = await Promise.all(Array.from({ length: 12 }, (_, i) => cliAsync("write", "b", `A${i + 1}`, `[[${i + 1}]]`)));
    expect(codes).toEqual(Array(12).fill(0));
    expect(cli("read", "b", "A1:A12").json().values.flat()).toEqual(Array.from({ length: 12 }, (_, i) => i + 1));
  });

  test("a deep formula is one cell error, and read and export still return the rest", () => {
    cli("new", "b");
    cli("write", "b", "A1", JSON.stringify([[1, `=${"(".repeat(4000)}1${")".repeat(4000)}`, 3]]));
    expect(cli("read", "b", "A1:C1").json().values).toEqual([[1, { error: "#DEPTH!" }, 3]]);
    expect(cli("export", "b").out).toBe("1,#DEPTH!,3\r\n");
    expect(cli("read", "b", "B1").json().formulas[0][0]).toMatch(/^=\(+1\)+$/);
  });

  test("a sheet of 150,000 rows imports and exports", { timeout: 240_000 }, () => {
    const file = join(dir, "big.csv");
    writeFileSync(file, Array.from({ length: 150_000 }, (_, i) => `${i + 1},row ${i + 1}\n`).join(""));
    cli("new", "b");
    expect(cli("import", "b", "--file", file).code).toBe(0);
    const lines = cli("export", "b").out.split("\r\n");
    expect([lines.length, lines[0], lines[149_999]]).toEqual([150_001, "1,row 1", "150000,row 150000"]);
  });
});

describe("sparse and wide sheets", () => {
  test("export refuses a sheet whose extent is too large, fast, and still serves a wide small one", { timeout: 20_000 }, () => {
    cli("new", "b");
    cli("write", "b", "A1", "[[1]]");
    cli("write", "b", "XFD1048576", "[[2]]");
    const started = Date.now();
    const r = cli("export", "b");
    expect([r.code, r.out]).toEqual([1, ""]);
    expect(r.err).toMatch(/^error: .*too large to export/);
    expect(r.err).not.toMatch(/\n\s+at /);
    expect(Date.now() - started).toBeLessThan(10_000);

    cli("new", "w");
    cli("write", "w", "A1", "[[1]]");
    cli("write", "w", "XFD3", "[[2]]");
    const lines = cli("export", "w").out.split("\r\n");
    expect([lines.length, lines[0].split(",").length, lines[2].endsWith(",2")]).toEqual([4, 16_384, true]);
  });
});

describe("csv text", () => {
  test("guarded text round-trips, a BOM goes, long ids and zeros stay text, --as-text keeps = as text", () => {
    cli("new", "b");
    const text = ["-5 apples", "+1", "@x", "\tx", "12345678901234567890", "007"];
    cli("write", "b", "A1", JSON.stringify([text]));
    const csv = cli("export", "b").out;
    expect(csv).toBe("'-5 apples,'+1,'@x,'\tx,12345678901234567890,007\r\n");
    const file = join(dir, "t.csv");
    writeFileSync(file, `\uFEFF${csv}`);
    cli("new", "c");
    cli("import", "c", "--file", file);
    expect(cli("read", "c", "A1:F1").json().values).toEqual([text]);
    writeFileSync(file, "=1+1,=2+2\n");
    cli("new", "d");
    cli("new", "e");
    cli("import", "d", "--file", file);
    cli("import", "e", "--file", file, "--as-text");
    expect(cli("read", "d", "A1:B1").json().values).toEqual([[2, 4]]);
    expect(cli("read", "e", "A1:B1").json().values).toEqual([["'=1+1", "'=2+2"]]);
  });
});

describe("mcp over stdio", () => {
  test("tools match the cli, label agent edits, and report bad input as errors", async () => {
    const client = new Client({ name: "test", version: "0" });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [CLI, "mcp"], env: { SHEET_DIR: dir, SHEET_AS: "bot", SHEET_KIND: "agent" } }));
    try {
      const names = (await client.listTools()).tools.map((t) => t.name).sort();
      expect(names).toEqual(["export", "formula", "import", "insert_rows", "list", "new", "read", "write"].map((n) => `sheet_${n}`));
      const call = async (name: string, args: Record<string, unknown>) => {
        const r = (await client.callTool({ name, arguments: args })) as { isError?: boolean; content: { text: string }[] };
        const { text } = r.content[0];
        return { isError: r.isError === true, body: /^[{[]/.test(text) ? JSON.parse(text) : text };
      };
      expect((await call("sheet_new", { book: "m" })).isError).toBe(false);
      await call("sheet_write", { book: "m", anchor: "A1", values: [[2, "=A1*3"]] });
      const read = (await call("sheet_read", { book: "m", range: "A1:B1" })).body;
      expect(read.values).toEqual([[2, 6]]);
      expect(read.authors[0][1]).toMatchObject({ kind: "agent", machine: true });
      expect(cli("read", "m", "A1:B1").json().values).toEqual([[2, 6]]);
      expect((await call("sheet_read", { book: "nope", range: "A1" })).isError).toBe(true);
      expect((await call("sheet_write", { book: "m", anchor: "A1", values: "x" })).isError).toBe(true);
      expect((await call("sheet_new", { book: "../x" })).isError).toBe(true);
      expect((await call("sheet_insert_rows", { book: "m", at: 1, count: 2_000_000 })).isError).toBe(true);
      await call("sheet_new", { book: "p" });
      await Promise.all(Array.from({ length: 12 }, (_, i) => call("sheet_write", { book: "p", anchor: `A${i + 1}`, values: [[i + 1]] })));
      expect(cli("read", "p", "A1:A12").json().values.flat()).toEqual(Array.from({ length: 12 }, (_, i) => i + 1));
    } finally {
      await client.close();
    }
  });
});
