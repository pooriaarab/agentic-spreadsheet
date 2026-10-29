import { createMemoryBlobs, createMemoryStore } from "agentic-identity";
import { describe, expect, test } from "vitest";
import { resourceId } from "./access.js";
import { trustHeaders } from "./auth.js";
import { createApp } from "./worker.js";

// Failure modes, listed before the code:
// - a non-owner invites, makes a link, lists, or revokes
// - an invite or a link gives owner, or has a bad role, lifetime, or both an email and a handle
// - a list shows an email, a token, or a token hash
// - an owner of two books revokes the link of one through the other
// - a revoked link is not marked, or a missing link looks like a success

const app = (store = createMemoryStore()) => createApp({ store, blobs: createMemoryBlobs(), authenticate: trustHeaders });
async function call(a: ReturnType<typeof app>, who: string, method: string, path: string, body?: unknown) {
  const res = await a.request(`/api/v1/sheets${path}`, {
    method,
    headers: { "x-principal-id": who, "x-principal-kind": "human", "content-type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  return { status: res.status, json: (await res.json().catch(() => ({}))) as Record<string, any> };
}

describe("invites and share links", () => {
  test("the owner invites, makes a link, lists both, and revokes the link", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "b" });
    const invite = await call(a, "ana", "POST", "/b/invites", { role: "viewer", email: "Ada@Example.com", expiresInMs: 60_000 });
    expect(invite.status).toBe(201);
    const link = await call(a, "ana", "POST", "/b/links", { role: "commenter", signedInOnly: true, expiresInMs: 60_000 });
    expect(link.status).toBe(201);
    const invites = await call(a, "ana", "GET", "/b/invites");
    expect(invites.json.invites).toMatchObject([{ id: invite.json.id, for: "email", role: "viewer" }]);
    expect(Object.keys(invites.json.invites[0]).sort()).toEqual(["expiresAt", "for", "id", "role"]);
    const links = await call(a, "ana", "GET", "/b/links");
    expect(links.json.links).toMatchObject([{ id: link.json.id, role: "commenter", signedInOnly: true }]);
    const text = JSON.stringify([invites.json, links.json]);
    for (const secret of ["example.com", link.json.token, "tokenHash"]) expect(text).not.toContain(secret);
    expect((await call(a, "ana", "DELETE", `/b/links/${link.json.id}`)).status).toBe(200);
    expect((await call(a, "ana", "GET", "/b/links")).json.links[0].revokedAt).toBeTypeOf("number");
  });

  test("a non-owner is refused every call", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "b" });
    await call(a, "ana", "PUT", "/b/share", { kind: "human", subject: "ed", role: "editor" });
    const link = await call(a, "ana", "POST", "/b/links", { role: "viewer" });
    for (const [method, path, body] of [["GET", "/b/invites"], ["POST", "/b/invites", { role: "viewer", handle: "x1" }], ["GET", "/b/links"], ["POST", "/b/links", { role: "viewer" }], ["DELETE", `/b/links/${link.json.id}`]] as const) {
      expect((await call(a, "ed", method, path, body)).status, path).toBe(403);
      expect((await call(a, "eve", method, path, body)).status, path).toBe(404);
    }
  });

  test("bad input is a 400", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "b" });
    for (const body of [{ role: "owner", handle: "bot" }, { role: "viewer" }, { role: "viewer", email: "a@b.co", handle: "bot" }, { role: "viewer", email: "nope" }, { role: "viewer", handle: "bot", expiresInMs: -1 }, { role: "viewer", handle: "bot", expiresInMs: "1" }]) {
      expect((await call(a, "ana", "POST", "/b/invites", body)).status, JSON.stringify(body)).toBe(400);
    }
    expect((await call(a, "ana", "POST", "/b/links", { role: "viewer", expiresInMs: 1e15 })).json.error).toContain("expiresInMs");
    for (const body of [{ role: "owner" }, { role: "viewer", signedInOnly: "yes" }, { role: "viewer", expiresInMs: 0 }, { role: "viewer", extra: 1 }]) {
      expect((await call(a, "ana", "POST", "/b/links", body)).status, JSON.stringify(body)).toBe(400);
    }
  });

  test("a link of another book is 404, and so is a missing link", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "one" });
    await call(a, "ana", "POST", "", { book: "two" });
    const link = await call(a, "ana", "POST", "/two/links", { role: "viewer" });
    expect((await call(a, "ana", "DELETE", `/one/links/${link.json.id}`)).status).toBe(404);
    expect((await call(a, "ana", "DELETE", "/one/links/missing")).status).toBe(404);
    expect((await call(a, "ana", "GET", "/two/links")).json.links[0].revokedAt).toBeUndefined();
  });

  test("the owner revokes an invite, and an invite of another book is 404", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "one" });
    await call(a, "ana", "POST", "", { book: "two" });
    const invite = await call(a, "ana", "POST", "/two/invites", { role: "viewer", handle: "helper" });
    expect((await call(a, "ana", "DELETE", `/one/invites/${invite.json.id}`)).status).toBe(404);
    expect((await call(a, "ana", "DELETE", "/one/invites/missing")).status).toBe(404);
    await call(a, "ana", "PUT", "/two/share", { kind: "human", subject: "ed", role: "editor" });
    expect((await call(a, "ed", "DELETE", `/two/invites/${invite.json.id}`)).status).toBe(403);
    expect((await call(a, "ana", "DELETE", `/two/invites/${invite.json.id}`)).status).toBe(200);
    expect((await call(a, "ana", "GET", "/two/invites")).json.invites[0].revokedAt).toBeTypeOf("number");
  });

  test("an accepted invite cannot be revoked, and a bad stored time never reads as open", async () => {
    const store = createMemoryStore();
    const a = app(store);
    await call(a, "ana", "POST", "", { book: "b" });
    const invite = await call(a, "ana", "POST", "/b/invites", { role: "viewer", handle: "helper" });
    const doc = (await store.get("invites", invite.json.id))!;
    await store.put("invites", invite.json.id, { ...doc, expiresAt: "soon" });
    expect((await call(a, "ana", "GET", "/b/invites")).json.invites).toMatchObject([{ id: invite.json.id, expiresAt: 0 }]);
    await store.put("invites", invite.json.id, { ...doc, acceptedAt: "yes" });
    expect((await call(a, "ana", "GET", "/b/invites")).json.invites).toMatchObject([{ id: invite.json.id, acceptedAt: 0 }]);
    expect((await call(a, "ana", "DELETE", `/b/invites/${invite.json.id}`)).status).toBe(400);
  });

  test("a stored link with a bad expiry lists as expired, not as never expiring", async () => {
    const store = createMemoryStore();
    const a = app(store);
    await call(a, "ana", "POST", "", { book: "b" });
    await store.put("links", "odd", { id: "odd", resourceId: await resourceId("b"), role: "viewer", tokenHash: "x", signedInOnly: false, expiresAt: "soon" });
    expect((await call(a, "ana", "GET", "/b/links")).json.links).toMatchObject([{ id: "odd", expiresAt: 0 }]);
  });

  test("a stored link with a role of owner is not listed", async () => {
    const store = createMemoryStore();
    const a = app(store);
    await call(a, "ana", "POST", "", { book: "b" });
    await store.put("links", "damaged", { id: "damaged", resourceId: await resourceId("b"), role: "owner", tokenHash: "x", signedInOnly: false });
    expect((await call(a, "ana", "GET", "/b/links")).json.links).toEqual([]);
  });

  test("a stored invite with a role of owner is not listed", async () => {
    const store = createMemoryStore();
    const a = app(store);
    await call(a, "ana", "POST", "", { book: "b" });
    await store.put("invites", "damaged", { id: "damaged", resourceId: await resourceId("b"), role: "owner", handle: "bot" });
    expect((await call(a, "ana", "GET", "/b/invites")).json.invites).toEqual([]);
  });

  test("revoking twice keeps the first revoke time", async () => {
    const a = app();
    await call(a, "ana", "POST", "", { book: "b" });
    const link = await call(a, "ana", "POST", "/b/links", { role: "viewer" });
    await call(a, "ana", "DELETE", `/b/links/${link.json.id}`);
    const first = (await call(a, "ana", "GET", "/b/links")).json.links[0].revokedAt;
    await new Promise((r) => setTimeout(r, 5));
    expect((await call(a, "ana", "DELETE", `/b/links/${link.json.id}`)).status).toBe(200);
    expect((await call(a, "ana", "GET", "/b/links")).json.links[0].revokedAt).toBe(first);
  });
});
