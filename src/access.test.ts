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
const as = (id: string, kind = "human") => ({ "x-principal-id": id, "x-principal-kind": kind });
async function call(a: ReturnType<typeof app>, who: string, method: string, path: string, body?: unknown) {
  const res = await a.request(`/api/v1/sheets${path}`, { method, headers: { ...as(who), "content-type": "application/json" }, body: body === undefined ? undefined : JSON.stringify(body) });
  return { status: res.status, json: (await res.json().catch(() => ({}))) as Record<string, unknown> };
}

describe("workbook access", () => {
  test("a stranger gets 404 on a book, as for a book that does not exist", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "plan.v1" });
    for (const [method, path, body] of [["GET", "/plan.v1/range?range=A1"], ["PUT", "/plan.v1/range", { anchor: "A1", values: [[1]] }], ["PUT", "/plan.v1/share", { subject: "eve", role: "editor" }]] as const) {
      const stranger = await call(a, "eve", method, path, body);
      const missing = await call(a, "eve", method, path.replace("plan.v1", "nothing"), body);
      expect([stranger.status, stranger.json]).toEqual([missing.status, { error: expect.stringContaining("no workbook") }]);
    }
  });

  test("roles: a viewer reads, an editor writes, only the owner shares, and a principal id with @ works", async () => {
    const a = app();
    await call(a, "ana@example.com", "POST", "", { book: "b" });
    expect((await call(a, "ana@example.com", "PUT", "/b/share", { subject: "vic", role: "viewer" })).status).toBe(200);
    expect((await call(a, "ana@example.com", "PUT", "/b/share", { subject: "ed", role: "editor" })).status).toBe(200);
    expect((await call(a, "vic", "GET", "/b/range?range=A1")).status).toBe(200);
    expect((await call(a, "vic", "PUT", "/b/range", { anchor: "A1", values: [[1]] })).status).toBe(403);
    expect((await call(a, "ed", "PUT", "/b/range", { anchor: "A1", values: [[1]] })).status).toBe(200);
    expect((await call(a, "ed", "PUT", "/b/share", { subject: "eve", role: "editor" })).status).toBe(403);
    expect((await call(a, "ana@example.com", "PUT", "/b/range", { anchor: "A2", values: [[2]] })).status).toBe(200);
  });

  test("a revoked principal loses access, and the share body is checked", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "b" });
    await call(a, "ana", "PUT", "/b/share", { subject: "ed", role: "editor" });
    expect((await call(a, "ana", "DELETE", "/b/share/ed")).status).toBe(200);
    expect((await call(a, "ed", "GET", "/b/range?range=A1")).status).toBe(404);
    for (const body of [{ subject: "x", role: "owner" }, { subject: "x", role: "root" }, { subject: "a b", role: "viewer" }, { subject: "x", role: "viewer", extra: 1 }]) {
      expect((await call(a, "ana", "PUT", "/b/share", body)).status).toBe(400);
    }
    // The last owner stays.
    expect((await call(a, "ana", "DELETE", "/b/share/ana")).status).toBe(400);
  });

  test("an agent that creates a book owns it", async () => {
    const a = app();
    expect((await call(a, "bot", "POST", "", { book: "mine" }, )).status).toBe(201);
    expect((await call(a, "bot", "PUT", "/mine/range", { anchor: "A1", values: [[1]] })).status).toBe(200);
  });
});
