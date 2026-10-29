// The adapters run the shared contract suites. Ways the workbook layer can fail on them:
//  1. Two callers on one store edit the same workbook, and one edit is lost or applied twice.
//  2. A failed edit leaves a state blob, or points a workbook at one.
//  3. A transaction that runs again leaves the blob of its first run.
//  4. A restart loses a saved edit.
import { existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createMemoryBlobs, createMemoryStore, type Blobs, type Store } from "agentic-identity";
import { runBlobsContract, runStoreContract } from "agentic-identity/testing";
import { z } from "zod";
import { afterAll, describe, expect, test } from "vitest";
import { createFileBlobs, createFileStore } from "./core/file-store.js";
import { OPS, type Ctx } from "./ops.js";

const dirs: string[] = [];
const fresh = () => dirs[dirs.push(mkdtempSync(join(tmpdir(), "sheet-ports-"))) - 1];
afterAll(() => dirs.forEach((d) => rmSync(d, { recursive: true, force: true })));
const shared = new Map<string, string>();
const dirFor = (name?: string) => (name === undefined ? fresh() : (shared.get(name) ?? shared.set(name, fresh()).get(name)!));

describe("memory adapters", () => {
  runStoreContract((options) => createMemoryStore(options));
  runBlobsContract((options) => createMemoryBlobs(options));
});
describe("file adapters", () => {
  runStoreContract((options) => createFileStore(dirFor(options.shared), options));
  runBlobsContract((options) => createFileBlobs(dirFor(options.shared), options));
});

const run = (name: string, args: object, ctx: Ctx) => OPS.find((o) => o.name === name)!.run(args, ctx);
const ctxOf = (store: Store, blobs: Blobs): Ctx => ({ store, blobs, principal: () => ({ id: "bot", kind: "agent" }) });
const keys = async (dir: string) => readdirSync(dir).filter((n) => !n.endsWith(".type") && !n.includes(".tmp"));

describe("workbooks on shared ports", () => {
  test("concurrent edits through two contexts on one store lose none and apply none twice", async () => {
    const [one, two] = [1, 2].map(() => ctxOf(createMemoryStore({ shared: "race" }), createMemoryBlobs({ shared: "race" })));
    await run("new", { book: "race" }, one);
    await Promise.all(Array.from({ length: 24 }, (_, i) => run("write", { book: "race", anchor: `A${i + 1}`, values: [[i]] }, i % 2 ? one : two)));
    await Promise.all([run("insert-rows", { book: "race", at: 1, count: 1 }, one), run("insert-rows", { book: "race", at: 1, count: 1 }, two)]);
    const read = (await run("read", { book: "race", range: "A1:A28" }, one)) as { values: unknown[][] };
    expect(read.values.flat().filter((v) => v !== null).sort((a, b) => Number(a) - Number(b))).toEqual(Array.from({ length: 24 }, (_, i) => i));
    expect(read.values.slice(0, 2)).toEqual([[null], [null]]);
  });

  test("two file stores on one directory lose none and apply none twice", async () => {
    const dir = fresh();
    const open = () => ctxOf(createFileStore(dir), createFileBlobs(join(dir, "blobs")));
    const [one, two] = [open(), open()];
    await run("new", { book: "race" }, one);
    await Promise.all(Array.from({ length: 16 }, (_, i) => run("insert-rows", { book: "race", at: 1, count: 1 }, i % 2 ? one : two)));
    await Promise.all(Array.from({ length: 16 }, (_, i) => run("write", { book: "race", anchor: `A${i + 17}`, values: [[i]] }, i % 2 ? one : two)));
    const read = (await run("read", { book: "race", range: "A1:A32" }, one)) as { values: unknown[][] };
    expect(read.values.slice(0, 16).flat()).toEqual(Array(16).fill(null));
    expect(read.values.slice(16).flat().sort((a, b) => Number(a) - Number(b))).toEqual(Array.from({ length: 16 }, (_, i) => i));
  });

  test("list returns every book when no limit is set, and a broken or odd old file is an error", async () => {
    const [store, blobs] = [createMemoryStore(), createMemoryBlobs()];
    const ctx = ctxOf(store, blobs);
    for (let i = 0; i < 130; i += 1) await store.put("workbooks", `b${String(i).padStart(3, "0")}`, { name: "x" });
    expect(((await run("list", {}, ctx)) as { books: string[] }).books).toHaveLength(130);
    const dir = fresh();
    mkdirSync(join(dir, "workbooks"), { recursive: true });
    writeFileSync(join(dir, "workbooks", "broken.json"), "{not json");
    writeFileSync(join(dir, "stray.json"), "{}");
    await expect(createFileStore(dir).get("workbooks", "broken")).rejects.toThrow(/not valid JSON/);
    await expect(createFileStore(dir).put("workbooks", "fresh", { name: "x" })).rejects.toThrow(/not valid JSON/);
    expect(existsSync(join(dir, "store.json"))).toBe(false);
    writeFileSync(join(dir, "workbooks", "broken.json"), "{}");
    writeFileSync(join(dir, "workbooks", "__x__.json"), "{}");
    await expect(createFileStore(dir).query("workbooks")).rejects.toThrow(/__x__\.json" has an id.*Rename/);
    const accepts = (book: string) => z.object(OPS.find((o) => o.name === "new")!.shape).safeParse({ book }).success;
    expect([accepts("__y__"), accepts("a".repeat(257)), accepts("_y_"), accepts("a".repeat(256))]).toEqual([false, false, true, true]);
  });

  test("a read creates no directory", async () => {
    const dir = join(fresh(), "none");
    const ctx = ctxOf(createFileStore(dir), createFileBlobs(join(dir, "blobs")));
    await expect(run("read", { book: "x", range: "A1" }, ctx)).rejects.toThrow(/no workbook/);
    expect(existsSync(dir)).toBe(false);
  });

  test("a failed edit leaves the old state and no extra blob", async () => {
    const dir = fresh();
    const ctx = ctxOf(createFileStore(dir), createFileBlobs(dir));
    await run("new", { book: "b" }, ctx);
    await run("write", { book: "b", anchor: "A1", values: [[1]] }, ctx);
    const before = (await keys(dir)).length;
    await expect(run("write", { book: "b", anchor: "A1", sheet: "Nope", values: [[2]] }, ctx)).rejects.toThrow();
    expect((await keys(dir)).length).toBe(before);
    expect(((await run("read", { book: "b", range: "A1:A1" }, ctx)) as { values: unknown[][] }).values).toEqual([[1]]);
    // One workbook state is one blob, and an edit replaces it.
    await run("write", { book: "b", anchor: "A1", values: [[3]] }, ctx);
    expect((await keys(dir)).length).toBe(before);
  });

  test("a transaction that runs again drops the blob of its first run", async () => {
    const dir = fresh();
    const blobs = createFileBlobs(dir);
    const store = createMemoryStore();
    await run("new", { book: "b" }, ctxOf(store, blobs));
    const count = () => readdirSync(dir).filter((n) => !n.endsWith(".type")).length;
    const before = count();
    let writes = 0;
    const racing: Store = {
      ...store,
      transaction: (fn) => store.transaction(async (tx) => {
        const result = await fn(tx);
        // Another writer changes the workbook document during the first run only.
        if (writes++ === 0) await store.put("workbooks", "b", { ...(await store.get("workbooks", "b"))!, updatedAt: "x" });
        return result;
      }),
    };
    await run("write", { book: "b", anchor: "A1", values: [[1]] }, ctxOf(racing, blobs));
    expect(writes).toBe(2);
    expect(count()).toBe(before);
  });

  test("edits survive a restart of the file adapters", async () => {
    const dir = fresh();
    const open = () => ctxOf(createFileStore(dir), createFileBlobs(join(dir, "blobs")));
    await run("new", { book: "b" }, open());
    await run("write", { book: "b", anchor: "B2", values: [["=1+2"]] }, open());
    expect(((await run("read", { book: "b", range: "B2:B2" }, open())) as { values: unknown[][] }).values).toEqual([[3]]);
  });
});
