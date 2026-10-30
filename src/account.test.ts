// Failure cases: shared owners lose data; kind collisions erase an agent; hidden cell stamps
// survive; damaged grants or blobs permit partial success; failed blob cleanup cannot retry.
import { createMemoryBlobs, createMemoryStore, createSharing, type Doc, type Query } from "agentic-identity";
import { expect, test } from "vitest";
import { createAccess, resourceId, subjectId } from "./access.js";
import { Workbook, type Principal } from "./core/workbook.js";
import { toBase64 } from "./core/bytes.js";
import { spreadsheetAccount } from "./account.js";

const ana: Principal = { id: "ana", kind: "human" };
const bot: Principal = { id: "ana", kind: "agent_of_human" };
const bob: Principal = { id: "bob", kind: "human" };
const fresh = () => ({ store: createMemoryStore(), blobs: createMemoryBlobs() });
async function make(deps: ReturnType<typeof fresh>, id: string, owner = ana, author = owner) {
  const wb = new Workbook();
  wb.addSheet("Sheet1");
  wb.writeRange("Sheet1", "A1", [["private", "kept"]], author);
  wb.setFormat("Sheet1", "B1", { bold: true }, bob);
  await deps.blobs.put(`workbooks/${id}/state`, wb.encode());
  await deps.store.put("workbooks", id, { name: id, blob: `workbooks/${id}/state`, updatedAt: new Date().toISOString() });
  await createAccess(deps.store).create(owner, id);
}

test("exports owned states, sharing, authored cells and machine labels as JSON", async () => {
  const deps = fresh();
  await make(deps, "mine");
  await make(deps, "theirs", bob, bot);
  const out = await spreadsheetAccount.exportAccount(ana, deps);
  expect(out.workbooks).toHaveLength(1);
  expect(out.workbooks[0]).toMatchObject({ id: "mine", state: expect.any(String) });
  expect(out.grants).toHaveLength(1);
  expect(out.authored).toEqual([]);
  const machine = await spreadsheetAccount.exportAccount(bot, deps);
  expect(machine.authored).toHaveLength(2);
  expect(machine.authored[0]).toMatchObject({ author: { kind: "agent_of_human", machine: true } });
  expect(JSON.parse(JSON.stringify(out))).toEqual(out);
});

test("deletes sole ownership and blobs, keeps co-owners and anonymises all kinds", async () => {
  const deps = fresh();
  await make(deps, "mine");
  await make(deps, "joint");
  const rid = await resourceId("joint");
  const aid = await subjectId(ana.id);
  await createSharing(deps.store).grant({ principalId: aid, kind: "human", chain: [{ id: aid, kind: "human" }], scopes: ["workbook:admin"] }, rid, { kind: "human", id: await subjectId(bob.id) }, "owner");
  await make(deps, "machine", bob, bot);
  await spreadsheetAccount.eraseAccount(ana, deps);
  expect(await deps.store.get("workbooks", "mine")).toBeNull();
  expect(await deps.blobs.get("workbooks/mine/state")).toBeNull();
  expect((await createAccess(deps.store).role(bob, "admin", "joint")).allowed).toBe(true);
  expect((await createAccess(deps.store).role(ana, "read", "joint")).allowed).toBe(false);
  await spreadsheetAccount.eraseAccount(bot, deps);
  const saved = await deps.store.get("workbooks", "machine");
  const bytes = (await deps.blobs.get(saved!.blob as string))!.data;
  expect(new TextDecoder().decode(bytes)).not.toContain('"ana"');
  const cell = Workbook.fromState(bytes).readRange("Sheet1", "A1:A1")[0]![0]!;
  expect(cell.author).toMatchObject({ id: "(erased)", kind: bot.kind, machine: true });
  expect(await spreadsheetAccount.eraseAccount(bot, deps)).toEqual({ workbooks: 0, grants: 0, authored: 0, invites: 0, links: 0, audit: 0 });
});

test("fails closed on damaged state and co-owner grants before deleting anything", async () => {
  const deps = fresh();
  await make(deps, "mine");
  await deps.store.put("workbooks", "bad", { state: toBase64(new Uint8Array([255])) });
  await expect(spreadsheetAccount.eraseAccount(ana, deps)).rejects.toThrow();
  expect(await deps.store.get("workbooks", "mine")).not.toBeNull();
  await deps.store.delete("workbooks", "bad");
  await deps.store.put("grants", "bad", { resourceId: await resourceId("mine"), subjectId: "other", role: "owner" });
  await expect(spreadsheetAccount.eraseAccount(ana, deps)).rejects.toThrow(/damaged/);
  expect(await deps.blobs.get("workbooks/mine/state")).not.toBeNull();
});

test("resumes after a blob delete fails and counts zero after completion", async () => {
  const deps = fresh();
  await make(deps, "mine");
  const remove = deps.blobs.delete.bind(deps.blobs);
  deps.blobs.delete = async () => { throw new Error("offline"); };
  await expect(spreadsheetAccount.eraseAccount(ana, deps)).rejects.toThrow();
  deps.blobs.delete = remove;
  await spreadsheetAccount.eraseAccount(ana, deps);
  expect(await deps.store.get("workbooks", "mine")).toBeNull();
  expect(await deps.store.get("resources", await resourceId("mine"))).toBeNull();
  expect(await spreadsheetAccount.eraseAccount(ana, deps)).toEqual({ workbooks: 0, grants: 0, authored: 0, invites: 0, links: 0, audit: 0 });
});

test("keeps formulas and format authors when rebuilding a shared workbook", async () => {
  const deps = fresh();
  await make(deps, "shared", bob, bot);
  const wb = Workbook.fromState((await deps.blobs.get("workbooks/shared/state"))!.data);
  wb.writeRange("Sheet1", "C1", [["=A1"]], bob);
  await deps.blobs.put("workbooks/shared/state", wb.encode());
  expect((await spreadsheetAccount.exportAccount(bob, deps)).workbooks).toHaveLength(1);
  await spreadsheetAccount.eraseAccount(bot, deps);
  const after = Workbook.fromState((await deps.blobs.get((await deps.store.get("workbooks", "shared"))!.blob as string))!.data);
  expect(after.readRange("Sheet1", "C1:C1")[0]![0]!.formula).toBe("=A1");
  expect(after.readRange("Sheet1", "B1:B1")[0]![0]!.formatAuthor).toMatchObject({ id: bob.id });
});

test("retries old-blob cleanup after a shared state replacement commits", async () => {
  const deps = fresh();
  await make(deps, "shared", bob, bot);
  const remove = deps.blobs.delete.bind(deps.blobs);
  deps.blobs.delete = async () => { throw new Error("offline"); };
  await expect(spreadsheetAccount.eraseAccount(bot, deps)).rejects.toThrow();
  expect((await deps.store.get("workbooks", "shared"))!.blob).not.toBe("workbooks/shared/state");
  deps.blobs.delete = remove;
  await spreadsheetAccount.eraseAccount(bot, deps);
  expect(await deps.blobs.get("workbooks/shared/state")).toBeNull();
  expect((await spreadsheetAccount.exportAccount(bot, deps)).authored).toEqual([]);
});

test("refuses invalid subjects and damaged stamps, including on unrelated workbooks", async () => {
  const deps = fresh();
  await expect(spreadsheetAccount.eraseAccount(ana, { ...deps, blobs: { ...deps.blobs, put: undefined } } as never)).rejects.toThrow();
  await expect(spreadsheetAccount.eraseAccount({ id: "", kind: "agent" }, deps)).rejects.toThrow();
  await make(deps, "shared", bob);
  const wb = Workbook.fromState((await deps.blobs.get("workbooks/shared/state"))!.data);
  const cells = wb.doc.getArray<import("yjs").Map<unknown>>("sheets").get(0)!.get("cells") as import("yjs").Map<unknown>;
  const key = [...cells.keys()].find((k) => k.endsWith(":by"))!;
  cells.set(key, { id: "ana", kind: "unknown", at: "bad" });
  await deps.blobs.put("workbooks/shared/state", wb.encode());
  await expect(spreadsheetAccount.exportAccount(ana, deps)).rejects.toThrow(/damaged/);
  await expect(spreadsheetAccount.eraseAccount(ana, deps)).rejects.toThrow(/damaged/);
});

test("retains an owner inherited from a folder", async () => {
  const deps = fresh();
  await make(deps, "mine");
  const rid = await resourceId("mine");
  await deps.store.put("resources", "folder", { id: "folder", type: "folder" });
  const hash = await subjectId(bob.id);
  await deps.store.put("grants", `folder:human:${hash}`, { resourceId: "folder", subjectKind: "human", subjectId: hash, role: "owner" });
  await deps.store.put("resources", rid, { ...(await deps.store.get("resources", rid))!, parentId: "folder" });
  expect(await spreadsheetAccount.eraseAccount(ana, deps)).toMatchObject({ workbooks: 0, grants: 1 });
  expect(await deps.store.get("workbooks", "mine")).not.toBeNull();
});

test("resumes when sharing cleanup fails after the workbook row goes", async () => {
  const deps = fresh();
  await make(deps, "mine");
  const transact = deps.store.transaction.bind(deps.store);
  let calls = 0;
  deps.store.transaction = async (fn) => { if (++calls === 2) throw new Error("offline"); return transact(fn); };
  await expect(spreadsheetAccount.eraseAccount(ana, deps)).rejects.toThrow();
  expect(await deps.store.get("workbooks", "mine")).toBeNull();
  deps.store.transaction = transact;
  await spreadsheetAccount.eraseAccount(bob, deps);
  expect(await deps.store.get("resources", await resourceId("mine"))).toBeNull();
});

test("redacts invite addresses and link secrets, and removes only the subject's metadata", async () => {
  const deps = fresh();
  await make(deps, "shared", bob, bot);
  const rid = await resourceId("shared"), hash = await subjectId(bot.id);
  await deps.store.put("invites", "mine", { resourceId: rid, createdBy: hash, email: "private@example.test", role: "viewer" });
  await deps.store.put("audit", "mine", { resourceId: rid, action: "invite.create", targetId: "mine", targetKind: "invite", principalId: hash, principalKind: bot.kind, rootId: hash, chain: [{ id: hash, kind: "agent" }] });
  await deps.store.put("links", "kept", { resourceId: rid, tokenHash: "secret", role: "viewer" });
  expect((await spreadsheetAccount.exportAccount(ana, deps)).invites).toEqual([]);
  const mine = await spreadsheetAccount.exportAccount(bot, deps);
  expect(mine.invites).toMatchObject([{ id: "mine", for: "email" }]);
  expect(JSON.stringify(mine)).not.toContain("private@example.test");
  expect(JSON.stringify(await spreadsheetAccount.exportAccount(bob, deps))).not.toContain("secret");
  expect(await spreadsheetAccount.eraseAccount(bot, deps)).toMatchObject({ invites: 1, audit: 1 });
  expect(await deps.store.get("links", "kept")).not.toBeNull();
});

test.each(["before", "after"])("candidate cleanup %s upload keeps live state and removes aborted bytes", async (when) => {
  const deps = fresh();
  await make(deps, "shared", bob, bot);
  const put = deps.blobs.put.bind(deps.blobs);
  let candidate = "";
  deps.blobs.put = async (key, data, options) => {
    candidate = key;
    if (when === "before") await spreadsheetAccount.eraseAccount({ id: "carol", kind: "human" }, deps);
    const result = await put(key, data, options);
    if (when === "after") await spreadsheetAccount.eraseAccount({ id: "carol", kind: "human" }, deps);
    return result;
  };
  await expect(spreadsheetAccount.eraseAccount(bot, deps)).rejects.toThrow(/retry/);
  expect(await deps.blobs.get((await deps.store.get("workbooks", "shared"))!.blob as string)).not.toBeNull();
  expect(await deps.blobs.get(candidate)).toBeNull();
  deps.blobs.put = put;
  await spreadsheetAccount.eraseAccount(bot, deps);
});

test("a damaged cleanup job cannot delete another owner's workbook", async () => {
  const deps = fresh();
  await make(deps, "shared", bob);
  await deps.store.put("spreadsheet_account_jobs", "bad", { owner: `human:${await subjectId(ana.id)}`, book: "shared", blob: "workbooks/shared/state", resource: await resourceId("shared") });
  await expect(spreadsheetAccount.eraseAccount(ana, deps)).rejects.toThrow();
  expect(await deps.blobs.get("workbooks/shared/state")).not.toBeNull();
});

test("format-only exports omit another author's content and identity", async () => {
  const deps = fresh();
  await make(deps, "mine", ana);
  const out = await spreadsheetAccount.exportAccount(bob, deps);
  expect(out.authored).toHaveLength(1);
  expect(out.authored[0]).toMatchObject({ address: "B1", shown: { format: { bold: true } } });
  expect(JSON.stringify(out.authored)).not.toContain('"ana"');
  expect(JSON.stringify(out.authored)).not.toContain("kept");
});

test("sole deletion captures a state another erase replaces after planning", async () => {
  const deps = fresh(), keys = new Set(["workbooks/mine/state"]);
  await make(deps, "mine", ana);
  const put = deps.blobs.put.bind(deps.blobs), transact = deps.store.transaction.bind(deps.store);
  deps.blobs.put = async (key, data, options) => { keys.add(key); return put(key, data, options); };
  let injected = false;
  deps.store.transaction = async (fn) => {
    if (!injected) { injected = true; await spreadsheetAccount.eraseAccount(bob, deps); }
    return transact(fn);
  };
  await spreadsheetAccount.eraseAccount(ana, deps);
  expect(await deps.store.get("workbooks", "mine")).toBeNull();
  for (const key of keys) expect(await deps.blobs.get(key)).toBeNull();
});

test("a deletion after the workbook listing returns a neutral retry error", async () => {
  const deps = fresh();
  await make(deps, "mine", ana);
  await make(deps, "bobs", bob);
  const query = deps.store.query.bind(deps.store);
  let injected = false;
  deps.store.query = async <T extends Doc = Doc>(collection: string, options?: Query) => {
    const page = await query<T>(collection, options);
    if (collection === "workbooks" && !injected) { injected = true; await spreadsheetAccount.eraseAccount(ana, deps); }
    return page;
  };
  await expect(spreadsheetAccount.eraseAccount(bob, deps)).rejects.toThrow(/^account scan changed; retry$/);
  deps.store.query = query;
  await spreadsheetAccount.eraseAccount(bob, deps);
});

test.each([true, false])("a held grant with a missing row distinguishes fence %s from damage", async (fenced) => {
  const deps = fresh();
  await make(deps, "mine", ana);
  await make(deps, "bobs", bob);
  await createAccess(deps.store).share(ana, "mine", { id: bob.id, kind: "human" }, "editor");
  const rid = await resourceId("mine"), owner = `human:${await subjectId(ana.id)}`, query = deps.store.query.bind(deps.store);
  let injected = false;
  deps.store.query = async <T extends Doc = Doc>(collection: string, options?: Query) => {
    if (collection === "workbooks" && !injected) {
      injected = true;
      await deps.store.delete("workbooks", "mine");
      await deps.store.put("resources", rid, { ...(await deps.store.get("resources", rid))!, spreadsheetErasing: fenced });
      if (fenced) await deps.store.put("spreadsheet_account_jobs", "pending", { owner, book: "mine", blob: "workbooks/mine/state", resource: rid });
    }
    return query<T>(collection, options);
  };
  await expect(spreadsheetAccount.exportAccount(bob, deps)).rejects.toThrow(fenced ? /^account scan changed; retry$/ : new RegExp(`damaged: resource ${rid}$`));
  deps.store.query = query;
  if (fenced) await spreadsheetAccount.eraseAccount(bob, deps);
});
