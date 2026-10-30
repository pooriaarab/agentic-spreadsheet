import * as Y from "yjs";
import { createSharing, ROLES, SUBJECT_KINDS, type Blobs, type Doc, type Json, type Store } from "agentic-identity";
import { resourceId, subjectId } from "./access.js";
import { fromBase64, toBase64 } from "./core/bytes.js";
import { Workbook, PRINCIPAL_KINDS, type Principal } from "./core/workbook.js";

export interface SpreadsheetAccountDeps { store: Store; blobs: Blobs }
export type SpreadsheetCollection = "workbooks" | "grants" | "authored" | "invites" | "links" | "audit";
/** Structural match for the identity account registry's AccountPackage contract. */
export interface SpreadsheetAccount {
  id: "spreadsheet";
  exportAccount(subject: Principal, deps: SpreadsheetAccountDeps): Promise<Record<SpreadsheetCollection, Json[]>>;
  eraseAccount(subject: Principal, deps: SpreadsheetAccountDeps): Promise<Record<SpreadsheetCollection, number>>;
}
export class SpreadsheetAccountError extends Error {}
const JOBS = "spreadsheet_account_jobs";
const ERASED = "(erased)";
function fail(): never { throw new SpreadsheetAccountError("account input or stored workbook records are damaged"); }
const plain = (value: unknown): Json => JSON.parse(JSON.stringify(value)) as Json;
async function rows(store: Store, collection: string, where?: Record<string, string>) {
  const out: { id: string; data: Doc }[] = [];
  let cursor: string | undefined;
  do {
    const page = await store.query(collection, { where, limit: 100, ...(cursor ? { cursor } : {}) });
    out.push(...page.items);
    cursor = page.nextCursor;
  } while (cursor !== undefined);
  return out;
}
function check(subject: Principal, deps: SpreadsheetAccountDeps) {
  if (typeof subject?.id !== "string" || !subject.id.trim() || subject.id === ERASED || !PRINCIPAL_KINDS.includes(subject.kind)) fail();
  if (!deps?.store?.transaction || !deps.blobs?.get || !deps.blobs?.delete) fail();
}
const matches = (value: unknown, subject: Principal) => {
  const stamp = value as Principal | null;
  return stamp?.id === subject.id && stamp?.kind === subject.kind;
};

/** Inspect every stamp, including cells in deleted rows. Damaged stamps cannot be skipped. */
function authored(wb: Workbook, subject: Principal, erase = false): Json[] {
  const found: Json[] = [];
  for (const sheet of wb.doc.getArray<Y.Map<unknown>>("sheets")) {
    if (!(sheet instanceof Y.Map) || typeof sheet.get("name") !== "string") fail();
    const cells = sheet.get("cells");
    if (!(cells instanceof Y.Map)) fail();
    for (const [key, value] of cells.entries()) {
      if (!key.endsWith(":by") && !key.endsWith(":fby")) continue;
      const stamp = value as Principal & { at?: unknown };
      if (!stamp || typeof stamp.id !== "string" || !PRINCIPAL_KINDS.includes(stamp.kind) || typeof stamp.at !== "string" || !Number.isFinite(Date.parse(stamp.at))) fail();
      if (!matches(value, subject)) continue;
      const cell = key.slice(0, key.lastIndexOf(":"));
      found.push(plain({ sheet: sheet.get("name"), cellId: cell, field: key.endsWith(":fby") ? "format" : "content", author: { ...stamp, machine: stamp.kind !== "human" }, entries: Object.fromEntries([...cells.entries()].filter(([k]) => k.startsWith(`${cell}:`))) }));
      if (erase) cells.set(key, { ...stamp, id: ERASED });
    }
  }
  return found;
}

/** Fresh CRDT structs discard replaced and deleted history, including old author IDs. */
function compact(wb: Workbook): string {
  const doc = new Y.Doc();
  doc.getArray("sheets").push(wb.doc.getArray<Y.Map<unknown>>("sheets").toArray().map((sheet) => sheet.clone()));
  return toBase64(Y.encodeStateAsUpdate(doc));
}
async function open(data: Doc, deps: SpreadsheetAccountDeps, bookId: string) {
  if (data.blob !== undefined && (typeof data.blob !== "string" || !data.blob)) fail();
  let state = data.state;
  if (data.blob !== undefined) {
    if (state !== undefined) fail();
    if (!(data.blob as string).startsWith(`workbooks/${bookId}/`)) fail();
    const blob = await deps.blobs.get(data.blob as string);
    if (!blob) fail();
    state = toBase64(blob.data);
  }
  if (typeof state !== "string") fail();
  try {
    const wb = Workbook.fromState(fromBase64(state));
    if (!wb.sheetNames().length) fail();
    return { wb, state };
  } catch { return fail(); }
}

async function plan(subject: Principal, deps: SpreadsheetAccountDeps) {
  check(subject, deps);
  const { store } = deps;
  const hash = await subjectId(subject.id);
  const kind = subject.kind === "human" ? "human" : "agent";
  const context = { principalId: hash, kind: subject.kind, chain: [{ id: hash, kind: kind as "human" | "agent" }], scopes: ["workbook:admin"] };
  const books = await rows(store, "workbooks");
  const names = new Map(await Promise.all(books.map(async (b) => [await resourceId(b.id), b.id] as const)));
  const grants: { id: string; data: Doc }[] = [];
  // The sharing layer collapses machine kinds to agent; the stamps above retain the exact kind.
  for (const grant of await rows(store, "grants", { subjectId: hash })) {
    if (!SUBJECT_KINDS.includes(grant.data.subjectKind as never)) fail();
    if (grant.data.subjectKind !== kind) continue;
    const { resourceId: rid, role } = grant.data;
    if (typeof rid !== "string" || !ROLES.includes(role as never)) fail();
    const resource = await store.get("resources", rid);
    if (resource?.type !== "workbook") continue;
    if (!names.has(rid)) fail();
    if (grant.id !== `${rid}:${kind}:${hash}`) fail();
    grants.push(grant);
  }
  const sole = new Set<string>();
  for (const grant of grants.filter((g) => g.data.role === "owner")) {
    let rid: string | undefined = grant.data.resourceId as string;
    let other = false;
    const seen = new Set<string>();
    while (rid !== undefined) {
      if (seen.has(rid) || seen.size >= 16) fail();
      seen.add(rid);
      const resource = await store.get("resources", rid);
      if (!resource) fail();
      for (const owner of await rows(store, "grants", { resourceId: rid })) {
        const g = owner.data;
        if (!ROLES.includes(g.role as never) || !SUBJECT_KINDS.includes(g.subjectKind as never) || typeof g.subjectId !== "string" || !g.subjectId || owner.id !== `${rid}:${String(g.subjectKind)}:${g.subjectId}`) fail();
        if (g.role === "owner" && (g.subjectKind !== kind || g.subjectId !== hash)) other = true;
      }
      if (resource.parentId != null && (typeof resource.parentId !== "string" || !resource.parentId)) fail();
      rid = resource.parentId as string | undefined;
      if (rid === null) rid = undefined;
    }
    if (!other) {
      const resource = grant.data.resourceId as string;
      if ((await rows(store, "resources", { parentId: resource })).length) fail();
      sole.add(names.get(resource)!);
    }
  }
  const opened = [];
  for (const book of books) {
    const openedBook = await open(book.data, deps, book.id);
    const stamps = authored(openedBook.wb, subject);
    opened.push({ ...book, ...openedBook, stamps, resource: await resourceId(book.id) });
  }
  const owned = new Set(grants.filter((g) => g.data.role === "owner").map((g) => g.data.resourceId));
  const metadata: { collection: "invites" | "links" | "audit"; id: string; data: Doc; mine: boolean }[] = [];
  for (const collection of ["invites", "links", "audit"] as const) {
    for (const row of await rows(store, collection)) {
      const rid = row.data.resourceId;
      const mentions = collection === "invites" ? row.data.createdBy === hash : collection === "audit" &&
        (row.data.principalId === hash || row.data.targetId === hash || row.data.rootId === hash || (Array.isArray(row.data.chain) && row.data.chain.some((v) => (v as { id?: unknown })?.id === hash)));
      if (!mentions && !owned.has(rid)) continue;
      if (typeof rid !== "string") fail();
      if ((await store.get("resources", rid))?.type !== "workbook") continue;
      let mine = collection === "invites" && mentions;
      if (collection === "audit" && mentions) {
        const d = row.data;
        if (d.principalId === hash && !PRINCIPAL_KINDS.includes(d.principalKind as never)) fail();
        if (d.targetId === hash && !SUBJECT_KINDS.includes(d.targetKind as never)) fail();
        if (d.rootId === hash && !Array.isArray(d.chain)) fail();
        const chain = Array.isArray(d.chain) ? d.chain : [];
        if (chain.some((v) => (v as { id?: unknown })?.id === hash && !SUBJECT_KINDS.includes((v as { kind: never }).kind))) fail();
        mine = (d.principalId === hash && d.principalKind === subject.kind) || (d.targetId === hash && d.targetKind === kind) || chain.some((v) => (v as { id?: unknown })?.id === hash && (v as { kind?: unknown }).kind === kind);
      }
      if (mine || owned.has(rid)) metadata.push({ collection, ...row, mine });
    }
  }
  return { grants, sole, opened, context, hash, kind, metadata };
}

/** Finish durable cleanup before discovering new work. A failed delete never loses its key. */
async function resume(subject: Principal, deps: SpreadsheetAccountDeps) {
  const owner = `${subject.kind}:${await subjectId(subject.id)}`;
  for (const job of await rows(deps.store, JOBS, { owner })) {
    const { book, blob, resource } = job.data;
    if (typeof book !== "string" || (blob !== undefined && typeof blob !== "string")) fail();
    if (blob !== undefined && !(blob as string).startsWith(`workbooks/${book}/`)) fail();
    if (resource !== undefined && resource !== await resourceId(book)) fail();
    const hash = await subjectId(subject.id);
    const kind = subject.kind === "human" ? "human" : "agent";
    const context = { principalId: hash, kind: subject.kind, chain: [{ id: hash, kind: kind as "human" | "agent" }], scopes: ["workbook:admin"] };
    if (blob) await deps.blobs.delete(blob as string);
    if (resource !== undefined) {
      await deps.store.delete("workbooks", book);
      if (await deps.store.get("resources", resource as string)) await createSharing(deps.store).deleteResource(context, resource as string);
    }
    await deps.store.delete(JOBS, job.id);
  }
}

export const spreadsheetAccount: SpreadsheetAccount = {
  id: "spreadsheet",
  async exportAccount(subject, deps) {
    const p = await plan(subject, deps);
    if ((await rows(deps.store, JOBS, { owner: `${subject.kind}:${p.hash}` })).length) throw new SpreadsheetAccountError("finish the interrupted erase before exporting");
    const owned = new Set(p.grants.filter((g) => g.data.role === "owner").map((g) => g.data.resourceId));
    return {
      workbooks: p.opened.filter((b) => owned.has(b.resource)).map((b) => plain({ id: b.id, state: b.state, updatedAt: b.data.updatedAt })),
      invites: p.metadata.filter((r) => r.collection === "invites").map((r) => plain({ id: r.id, ...r.data })),
      links: p.metadata.filter((r) => r.collection === "links").map((r) => { const { tokenHash: _secret, ...data } = r.data; return plain({ id: r.id, ...data }); }),
      audit: p.metadata.filter((r) => r.collection === "audit").map((r) => plain({ id: r.id, ...r.data })),
      grants: p.grants.map((g) => plain({ id: g.id, ...g.data })),
      authored: p.opened.filter((b) => !owned.has(b.resource)).flatMap((b) => b.stamps.map((stamp) => plain({ workbookId: b.id, ...stamp as object }))),
    };
  },
  async eraseAccount(subject, deps) {
    check(subject, deps);
    await resume(subject, deps);
    const p = await plan(subject, deps);
    const counts = { workbooks: 0, grants: p.grants.length, authored: 0, invites: 0, links: 0, audit: 0 };
    const owner = `${subject.kind}:${p.hash}`;
    for (const b of p.opened) {
      if (p.sole.has(b.id)) {
        await deps.store.put(JOBS, crypto.randomUUID(), { owner, book: b.id, ...(b.data.blob ? { blob: b.data.blob } : {}), resource: await resourceId(b.id) });
        await resume(subject, deps);
        counts.workbooks++;
      } else if (b.stamps.length) {
        authored(b.wb, subject, true);
        const state = compact(b.wb);
        // The replacement and old-blob cleanup key commit together. Stop writes while this runs.
        await deps.store.transaction(async (tx) => {
          await tx.put("workbooks", b.id, { name: b.id, updatedAt: b.data.updatedAt, state });
          if (b.data.blob) await tx.put(JOBS, crypto.randomUUID(), { owner, book: b.id, blob: b.data.blob });
        });
        await resume(subject, deps);
        counts.authored += b.stamps.length;
      }
    }
    for (const row of p.metadata) {
      if (row.mine) await deps.store.delete(row.collection, row.id);
      if (row.mine || p.sole.has(p.opened.find((b) => b.resource === row.data.resourceId)?.id ?? "")) counts[row.collection]++;
    }
    for (const grant of p.grants) await deps.store.delete("grants", grant.id);
    return counts;
  },
};
