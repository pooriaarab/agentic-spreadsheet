import { createMemoryBlobs, createMemoryStore } from "agentic-identity";
import { describe, expect, test } from "vitest";
import { createApp } from "./worker.js";
import { trustHeaders } from "./auth.js";

// Failure modes, listed before the code:
// - a principal with no role reads, writes, or shares a book, or learns that the book exists
// - a viewer writes, or an editor shares
// - a book id with a dot, or a principal id with "@", breaks the owner or a grant
// - a revoked principal keeps its access
// - the last owner is removed, or a share names a role of "owner" or an unknown role

const app = () => createApp({ store: createMemoryStore(), blobs: createMemoryBlobs(), authenticate: trustHeaders });
const as = (id: string, kind: string) => ({ "x-principal-id": id, "x-principal-kind": kind });
async function call(a: ReturnType<typeof app>, who: string, method: string, path: string, body?: unknown, kind = who === "bot" ? "agent" : "human") {
  const res = await a.request(`/api/v1/sheets${path}`, { method, headers: { ...as(who, kind), "content-type": "application/json" }, body: body === undefined ? undefined : JSON.stringify(body) });
  return { status: res.status, json: (await res.json().catch(() => ({}))) as Record<string, unknown> };
}

describe("workbook access", () => {
  test("a stranger gets 404 on a book, as for a book that does not exist", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "plan.v1" });
    for (const [method, path, body] of [["GET", "/plan.v1/range?range=A1"], ["PUT", "/plan.v1/range", { anchor: "A1", values: [[1]] }], ["PUT", "/plan.v1/share", { kind: "human", subject: "eve", role: "editor" }]] as const) {
      const stranger = await call(a, "eve", method, path, body);
      const missing = await call(a, "eve", method, path.replace("plan.v1", "nothing"), body);
      expect([stranger.status, stranger.json]).toEqual([missing.status, { error: expect.stringContaining("no workbook") }]);
    }
  });

  test("roles: a viewer reads, an editor writes, only the owner shares, and a principal id with @ works", async () => {
    const a = app();
    await call(a, "ana@example.com", "POST", "", { book: "b" });
    expect((await call(a, "ana@example.com", "PUT", "/b/share", { kind: "human", subject: "vic", role: "viewer" })).status).toBe(200);
    expect((await call(a, "ana@example.com", "PUT", "/b/share", { kind: "human", subject: "ed", role: "editor" })).status).toBe(200);
    expect((await call(a, "vic", "GET", "/b/range?range=A1")).status).toBe(200);
    expect((await call(a, "vic", "PUT", "/b/range", { anchor: "A1", values: [[1]] })).status).toBe(403);
    expect((await call(a, "ed", "PUT", "/b/range", { anchor: "A1", values: [[1]] })).status).toBe(200);
    expect((await call(a, "ed", "PUT", "/b/share", { kind: "human", subject: "eve", role: "editor" })).status).toBe(403);
    expect((await call(a, "ana@example.com", "PUT", "/b/range", { anchor: "A2", values: [[2]] })).status).toBe(200);
  });

  test("a revoked principal loses access, and the share body is checked", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "b" });
    await call(a, "ana", "PUT", "/b/share", { kind: "human", subject: "ed", role: "editor" });
    expect((await call(a, "ana", "DELETE", "/b/share/human/ed")).status).toBe(200);
    expect((await call(a, "ed", "GET", "/b/range?range=A1")).status).toBe(404);
    for (const body of [{ kind: "human", subject: "x", role: "owner" }, { kind: "human", subject: "x", role: "root" }, { kind: "human", subject: "a b", role: "viewer" }, { kind: "human", subject: "x", role: "viewer", extra: 1 }]) {
      expect((await call(a, "ana", "PUT", "/b/share", body)).status).toBe(400);
    }
    // The last owner stays.
    expect((await call(a, "ana", "DELETE", "/b/share/human/ana")).status).toBe(400);
  });

  test("the owner cannot lower itself, a commenter only reads, and a non-owner cannot unshare", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "b" });
    expect((await call(a, "ana", "PUT", "/b/share", { kind: "human", subject: "ana", role: "viewer" })).status).toBe(400);
    expect((await call(a, "ana", "PUT", "/b/range", { anchor: "A1", values: [[1]] })).status).toBe(200);
    await call(a, "ana", "PUT", "/b/share", { kind: "human", subject: "cy", role: "commenter" });
    expect((await call(a, "cy", "GET", "/b/range?range=A1")).status).toBe(200);
    expect((await call(a, "cy", "PUT", "/b/range", { anchor: "A1", values: [[2]] })).status).toBe(403);
    expect((await call(a, "cy", "DELETE", "/b/share/human/ana")).status).toBe(403);
    expect((await call(a, "eve", "DELETE", "/b/share/human/cy")).status).toBe(404);
    // A share that is not there is an error, not a success.
    expect((await call(a, "ana", "DELETE", "/b/share/human/nobody")).status).toBe(404);
  });

  test("a person and an agent with the same id hold separate roles", async () => {
    const a = app();
    await call(a, "kim", "POST", "", { book: "b" });
    const asAgent = (method: string, path: string, body?: unknown) => call(a, "kim", method, path, body, "agent");
    // The agent kim has no role on the person kim's book.
    expect((await asAgent("GET", "/b/range?range=A1")).status).toBe(404);
    expect((await call(a, "kim", "PUT", "/b/share", { kind: "agent", subject: "kim", role: "viewer" })).status).toBe(200);
    expect((await asAgent("GET", "/b/range?range=A1")).status).toBe(200);
    expect((await asAgent("PUT", "/b/range", { anchor: "A1", values: [[1]] })).status).toBe(403);
    // The person keeps the owner role, and the agent share does not touch it.
    expect((await call(a, "kim", "PUT", "/b/range", { anchor: "A1", values: [[1]] })).status).toBe(200);
    expect((await call(a, "kim", "DELETE", "/b/share/agent/kim")).status).toBe(200);
    expect((await asAgent("GET", "/b/range?range=A1")).status).toBe(404);
    expect((await call(a, "kim", "GET", "/b/range?range=A1")).status).toBe(200);
    // A share of kind human does not reach an agent with that id.
    expect((await call(a, "kim", "PUT", "/b/share", { kind: "human", subject: "lee", role: "editor" })).status).toBe(200);
    expect((await call(a, "lee", "GET", "/b/range?range=A1", undefined, "agent")).status).toBe(404);
  });

  test("an authenticator that returns an unknown kind gets no access", async () => {
    const s = createMemoryStore();
    const good = createApp({ store: s, blobs: createMemoryBlobs(), authenticate: trustHeaders });
    expect((await call(good, "bot", "POST", "", { book: "b" }, "agent")).status).toBe(201);
    const bad = createApp({ store: s, blobs: createMemoryBlobs(), authenticate: async () => ({ id: "bot", kind: "robot" }) as never });
    for (const method of ["GET", "DELETE"]) {
      const res = await bad.request(`/api/v1/sheets/b${method === "GET" ? "/range?range=A1" : ""}`, { method });
      expect(res.status).toBeGreaterThanOrEqual(400);
    }
    expect((await call(good, "bot", "GET", "/b/range?range=A1", undefined, "agent")).status).toBe(200);
  });

  test("deleting a book erases its grants, so a new book of that name starts clean", async () => {
    const s = createMemoryStore();
    const a = createApp({ store: s, blobs: createMemoryBlobs(), authenticate: trustHeaders });
    await call(a, "ana", "POST", "", { book: "b" });
    await call(a, "ana", "PUT", "/b/share", { kind: "human", subject: "ed", role: "editor" });
    expect((await call(a, "ed", "DELETE", "/b")).status).toBe(403);
    expect((await call(a, "eve", "DELETE", "/b")).status).toBe(404);
    expect((await call(a, "ana", "DELETE", "/b")).status).toBe(200);
    expect((await call(a, "ana", "GET", "/b/range?range=A1")).status).toBe(404);
    expect((await s.query("grants", { limit: 10 })).items).toEqual([]);
    expect((await s.query("workbooks", { limit: 10 })).items).toEqual([]);
    // Ed lost the editor role with the book.
    await call(a, "ana", "POST", "", { book: "b" });
    expect((await call(a, "ed", "GET", "/b/range?range=A1")).status).toBe(404);
  });

  test("a failed blob delete answers 500 and leaves the book, and the owner can delete it again", async () => {
    const [store, real] = [createMemoryStore(), createMemoryBlobs()];
    let fail = true;
    const blobs = { ...real, delete: async (key: string) => (fail ? Promise.reject(new Error("blob store down")) : real.delete(key)) };
    const a = createApp({ store, blobs, authenticate: trustHeaders });
    await call(a, "ana", "POST", "", { book: "b" });
    const blob = (await store.get("workbooks", "b"))!.blob as string;
    const failed = await call(a, "ana", "DELETE", "/b");
    expect([failed.status, JSON.stringify(failed.json)]).toEqual([500, expect.not.stringContaining("blob store down")]);
    // The book is still there, and so are its grants, so the owner can call again.
    expect((await store.get("workbooks", "b")) !== null && (await store.query("grants", { limit: 10 })).items.length > 0).toBe(true);
    fail = false;
    expect((await call(a, "ana", "DELETE", "/b")).status).toBe(200);
    expect((await store.query("workbooks", { limit: 10 })).items).toEqual([]);
    expect(await real.get(blob)).toBeNull();
    expect((await store.query("grants", { limit: 10 })).items).toEqual([]);
  });

  test("a save that lands during a DELETE does not leave a book with no owner", async () => {
    const [store, real] = [createMemoryStore(), createMemoryBlobs()];
    let hook = true;
    const a = createApp({ store, blobs: { ...real, delete: async (key: string) => { if (hook) { hook = false; await call(a, "ana", "PUT", "/b/range", { anchor: "A1", values: [[1]] }); } return real.delete(key); } }, authenticate: trustHeaders });
    await call(a, "ana", "POST", "", { book: "b" });
    expect((await call(a, "ana", "DELETE", "/b")).status).toBe(200);
    expect((await store.query("workbooks", { limit: 10 })).items).toEqual([]);
    expect((await store.query("grants", { limit: 10 })).items).toEqual([]);
  });

  test("guard: a failed erase of the grants leaves the owner able to delete again, and no one can take the name meanwhile", async () => {
    const real = createMemoryStore();
    let fail = true;
    const store = { ...real, delete: async (collection: string, id: string) => (fail && collection === "grants" ? Promise.reject(new Error("store down")) : real.delete(collection, id)) };
    const a = createApp({ store, blobs: createMemoryBlobs(), authenticate: trustHeaders });
    await call(a, "ana", "POST", "", { book: "b" });
    await call(a, "ana", "PUT", "/b/share", { kind: "human", subject: "ed", role: "editor" });
    expect((await call(a, "ana", "DELETE", "/b")).status).toBeGreaterThanOrEqual(400);
    expect((await call(a, "eve", "POST", "", { book: "b" })).status).toBe(409);
    fail = false;
    expect((await call(a, "ana", "DELETE", "/b")).status).toBe(200);
    expect((await real.query("workbooks", { limit: 10 })).items).toEqual([]);
    expect((await real.query("grants", { limit: 10 })).items).toEqual([]);
    expect((await call(a, "eve", "POST", "", { book: "b" })).status).toBe(201);
    expect((await call(a, "ed", "GET", "/b/range?range=A1")).status).toBe(404);
  });

  test("an agent that creates a book owns it", async () => {
    const a = app();
    expect((await call(a, "bot", "POST", "", { book: "mine" })).status).toBe(201);
    expect((await call(a, "bot", "PUT", "/mine/range", { anchor: "A1", values: [[1]] })).status).toBe(200);
  });
});
