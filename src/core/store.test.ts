import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, test } from "vitest";
import { FileStore } from "./store.js";

// Failure modes, listed before the code:
// - two stores on one directory both read the same counter and one increment is lost
// - a lock left by a dead process, or by a live one that stopped renewing it, blocks every write
// - a live, fresh lock is taken over, so two writers run at once
// - a lock file written half-way (empty) is read as stale
// - two writers share one temp file name and interleave
// - an unknown or unsafe cursor restarts the list at page 1
// - the lock is not released after a callback throws, or a late release removes another holder's lock

let dir: string;
beforeEach(() => void (dir = mkdtempSync(join(tmpdir(), "store-"))));
afterEach(() => rmSync(dir, { recursive: true, force: true }));

const bump = (store: FileStore) =>
  store.transaction(async (tx) => {
    const n = ((await tx.get("c", "n"))?.n as number | undefined) ?? 0;
    await new Promise((r) => setTimeout(r, 1));
    await tx.put("c", "n", { n: n + 1 });
  });

describe("transactions", () => {
  test("two stores on one directory lose no update", async () => {
    const [a, b] = [new FileStore(dir), new FileStore(dir)];
    await Promise.all(Array.from({ length: 40 }, (_, i) => bump(i % 2 ? a : b)));
    expect((await a.get("c", "n"))?.n).toBe(40);
  });

  test("the lock is released after a throw, and a nested transaction does not wait on itself", async () => {
    const store = new FileStore(dir);
    await expect(store.transaction(async () => { throw new Error("boom"); })).rejects.toThrow("boom");
    await store.transaction((tx) => tx.transaction((inner) => inner.put("c", "a", { ok: true })));
    expect(await store.get("c", "a")).toEqual({ ok: true });
  });
});

describe("stale locks", () => {
  const lock = (body: object | string) => writeFileSync(join(dir, ".lock"), typeof body === "string" ? body : JSON.stringify(body));

  test("a lock from a dead process is taken over", async () => {
    lock({ pid: 2 ** 22 + 12345, token: "dead" });
    await bump(new FileStore(dir));
    expect((await new FileStore(dir).get("c", "n"))?.n).toBe(1);
  });

  test("a lock from a live process that stopped renewing it is taken over; a fresh one is not", async () => {
    lock({ pid: process.pid, token: "live" });
    await new Promise((r) => setTimeout(r, 120));
    await bump(new FileStore(dir, { staleMs: 100 }));
    lock({ pid: process.pid, token: "fresh" });
    await expect(new FileStore(dir, { staleMs: 60_000, waitMs: 150 }).transaction(async () => 1)).rejects.toThrow(/locked/);
  });

  test("a lock file with no content yet is not stale", async () => {
    lock("");
    await expect(new FileStore(dir, { staleMs: 60_000, waitMs: 150 }).transaction(async () => 1)).rejects.toThrow(/locked/);
  });
});

describe("query", () => {
  test("pages by id, and a bad cursor throws", async () => {
    const store = new FileStore(dir);
    for (const id of ["c", "a", "b"]) await store.put("w", id, { id });
    expect((await store.query("w", { limit: 2 })).items.map((i) => i.id)).toEqual(["a", "b"]);
    expect((await store.query("w", { limit: 2, cursor: "b" })).items.map((i) => i.id)).toEqual(["c"]);
    await expect(store.query("w", { cursor: "../x" })).rejects.toThrow(/cursor/);
  });
});
