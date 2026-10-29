import { createSharing, type PrincipalContext, type Role, type Store, type Subject, type SubjectKind } from "agentic-identity";
import { PRINCIPAL_KINDS, type Principal } from "./core/workbook.js";

// Who may use a workbook. agentic-identity keeps the roles. A workbook is a resource of type
// "workbook", and the principal that makes it is its owner.
const TYPE = "workbook";
const SCOPES = [`${TYPE}:read`, `${TYPE}:write`, `${TYPE}:admin`];
export class NoShare extends Error {
  constructor(subject: string) {
    super(`"${subject}" has no share on this workbook`);
  }
}

export class NotOwner extends Error {
  constructor() {
    super("only an owner may list the links of a workbook");
  }
}

export class NoLink extends Error {
  constructor(link: string) {
    super(`no link "${link}" on this workbook`);
  }
}

export interface LinkSummary {
  id: string;
  role: Role;
  signedInOnly: boolean;
  expiresAt?: number;
  revokedAt?: number;
}

export const SHARE_ROLES = ["viewer", "commenter", "editor"] as const;
export type ShareRole = (typeof SHARE_ROLES)[number];

async function hash(prefix: string, text: string): Promise<string> {
  const digest = new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text)));
  return prefix + [...digest].map((b) => b.toString(16).padStart(2, "0")).join("");
}

// Resource and subject ids in agentic-identity allow letters, digits, "_" and "-". Workbook ids and
// principal ids allow more, so both go in as a hash. The hash is not keyed: a short name can be guessed from it.
export const resourceId = (book: string) => hash("wb", book);
export const subjectId = (principal: string) => hash("p", principal);

// A role belongs to a person or to an agent. A person and an agent with one id hold two roles.
// Every kind but "human" is a machine, so it is an agent subject.
export const SUBJECT_KIND_NAMES = ["human", "agent"] as const;
function subjectKind(kind: Principal["kind"]): SubjectKind {
  if (kind === "human") return "human";
  // An unknown kind is refused, never read as an agent.
  if (PRINCIPAL_KINDS.includes(kind)) return "agent";
  throw new Error(`unknown principal kind "${String(kind)}"`);
}

async function context(who: Principal): Promise<PrincipalContext> {
  const id = await subjectId(who.id);
  return { principalId: id, kind: who.kind, chain: [{ id, kind: subjectKind(who.kind) }], scopes: SCOPES };
}

async function subject(kind: SubjectKind, id: string): Promise<Subject> {
  return { kind, id: await subjectId(id) };
}

// The best role the principal has on the book, or null when it has none. An unknown book gives null.
export type Roles = ReturnType<typeof createAccess>;
export function createAccess(store: Store) {
  const sharing = createSharing(store);
  return {
    async role(who: Principal, action: "read" | "write" | "admin", book: string): Promise<{ allowed: boolean; role: Role | null }> {
      const decision = await sharing.can(await context(who), action, await resourceId(book));
      return { allowed: decision.allowed, role: decision.role ?? null };
    },
    async create(who: Principal, book: string): Promise<void> {
      await sharing.createResource(await context(who), { id: await resourceId(book), type: TYPE });
    },
    // Remove the book from the sharing store with its grants, invites, links, and audit entries. Only an owner may.
    async erase(who: Principal, book: string): Promise<void> {
      await sharing.deleteResource(await context(who), await resourceId(book));
    },
    async share(who: Principal, book: string, to: { kind: SubjectKind; id: string }, role: ShareRole): Promise<void> {
      const [resource, target] = [await resourceId(book), await subject(to.kind, to.id)];
      // A grant overwrites the old one. Without this check the owner could lower itself and leave the book with no owner.
      if ((await store.get("grants", `${resource}:${target.kind}:${target.id}`))?.role === "owner") throw new Error("an owner keeps the owner role");
      await sharing.grant(await context(who), resource, target, role);
    },
    // The grants on a book, with hashed subject ids: the store never holds a principal id. Only an owner may.
    async list(who: Principal, book: string, paging: { limit?: number; cursor?: string } = {}): Promise<{ grants: { subjectKind: SubjectKind; subjectHash: string; role: Role }[]; nextCursor?: string }> {
      const page = await sharing.listGrants(await context(who), await resourceId(book), { limit: paging.limit ?? 100, cursor: paging.cursor });
      return { grants: page.grants.map((g) => ({ subjectKind: g.subjectKind, subjectHash: g.subjectId, role: g.role })), nextCursor: page.nextCursor };
    },
    // The invites on a book. An invite shows no email: it says whether it is for an email or a handle. Only an owner may.
    async invites(who: Principal, book: string, paging: { limit?: number; cursor?: string } = {}) {
      return sharing.listInvites(await context(who), await resourceId(book), { limit: paging.limit ?? 100, cursor: paging.cursor });
    },
    // Invites a person by email or an agent by handle. The role is not owner. Only an owner may.
    async invite(who: Principal, book: string, input: { role: ShareRole; email?: string; handle?: string; expiresInMs?: number }): Promise<{ id: string }> {
      return sharing.createInvite(await context(who), await resourceId(book), input);
    },
    // The links on a book, one page. The store keeps a token hash only, and a list never shows it. Only an owner may.
    async links(who: Principal, book: string, paging: { limit?: number; cursor?: string } = {}): Promise<{ links: LinkSummary[]; nextCursor?: string }> {
      const resource = await resourceId(book);
      if (!(await sharing.can(await context(who), "admin", resource)).allowed) throw new NotOwner();
      const page = await store.query("links", { where: { resourceId: resource }, limit: paging.limit ?? 100, ...(paging.cursor === undefined ? {} : { cursor: paging.cursor }) });
      // A doc with a bad id or a role that is not a share role is not honored, so it is not listed.
      const valid = page.items.filter(({ data: d }) => typeof d.id === "string" && (SHARE_ROLES as readonly unknown[]).includes(d.role));
      const links = valid.map(({ data: d }): LinkSummary => ({
        id: String(d.id),
        role: d.role as Role,
        signedInOnly: d.signedInOnly === true,
        ...(typeof d.expiresAt === "number" ? { expiresAt: d.expiresAt } : {}),
        ...(d.revokedAt != null ? { revokedAt: typeof d.revokedAt === "number" ? d.revokedAt : 0 } : {}),
      }));
      return { links, ...(page.nextCursor === undefined ? {} : { nextCursor: page.nextCursor }) };
    },
    // The token shows once, in this reply. The role is not owner. Only an owner may.
    async createLink(who: Principal, book: string, input: { role: ShareRole; signedInOnly?: boolean; expiresInMs?: number }, now = Date.now()): Promise<{ id: string; token: string }> {
      const { expiresInMs, ...rest } = input;
      return sharing.createLink(await context(who), await resourceId(book), { ...rest, ...(expiresInMs === undefined ? {} : { expiresAt: now + expiresInMs }) });
    },
    // A link of another book is "no such link" here, even for an owner of both.
    async revokeLink(who: Principal, book: string, linkId: string): Promise<void> {
      const resource = await resourceId(book);
      if (!/^[\w-]{1,128}$/.test(linkId) || (await store.get("links", linkId))?.resourceId !== resource) throw new NoLink(linkId);
      // Revoking twice keeps the first revoke time.
      if ((await store.get("links", linkId))?.revokedAt != null) return;
      await sharing.revokeLink(await context(who), linkId);
    },
    async unshare(who: Principal, book: string, from: { kind: SubjectKind; id: string }): Promise<void> {
      const [resource, target] = [await resourceId(book), await subject(from.kind, from.id)];
      // Revoking a share that does not exist must not look like success.
      if ((await store.get("grants", `${resource}:${target.kind}:${target.id}`)) === null) throw new NoShare(from.id);
      await sharing.revokeGrant(await context(who), resource, target);
    },
  };
}
