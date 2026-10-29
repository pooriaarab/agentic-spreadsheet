import { createSharing, type PrincipalContext, type Role, type Store } from "agentic-identity";
import type { Principal } from "./core/workbook.js";

// Who may use a workbook. agentic-identity keeps the roles. A workbook is a resource of type
// "workbook", and the principal that makes it is its owner.
const TYPE = "workbook";
const SCOPES = [`${TYPE}:read`, `${TYPE}:write`, `${TYPE}:admin`];
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

async function context(who: Principal): Promise<PrincipalContext> {
  const id = await subjectId(who.id);
  return { principalId: id, kind: who.kind, chain: [{ id, kind: who.kind }], scopes: SCOPES };
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
    async share(who: Principal, book: string, subject: string, role: ShareRole): Promise<void> {
      const [resource, to] = [await resourceId(book), await subjectId(subject)];
      // A grant overwrites the old one. Without this check the owner could lower itself and leave the book with no owner.
      if ((await store.get("grants", `${resource}:${to}`))?.role === "owner") throw new Error("an owner keeps the owner role");
      await sharing.grant(await context(who), resource, to, role);
    },
    async unshare(who: Principal, book: string, subject: string): Promise<void> {
      await sharing.revokeGrant(await context(who), await resourceId(book), await subjectId(subject));
    },
  };
}
