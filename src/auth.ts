import { createStoreAuth, isScope, type Auth, type Principal as IdentityPrincipal, type Store } from "agentic-identity";
import { PRINCIPAL_KINDS, type Principal, type PrincipalKind } from "./core/workbook.js";

// Who is calling. It returns null when the request has no valid credential. `scopes` is set when the credential limits what it may do.
export type Authenticator = (request: Request) => Promise<(Principal & { scopes?: string[] }) | null>;

export class ConfigError extends Error {}

const MIN_KEY_BYTES = 32;
const ID = /^[\w.@:-]{1,128}$/;
const encoder = new TextEncoder();

async function digestHex(text: string): Promise<string> {
  const hash = new Uint8Array(await crypto.subtle.digest("SHA-256", encoder.encode(text)));
  return [...hash].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function principal(id: unknown, kind: unknown): Principal | null {
  if (typeof id !== "string" || !ID.test(id) || !PRINCIPAL_KINDS.includes(kind as PrincipalKind)) return null;
  return { id, kind: kind as PrincipalKind };
}

// Who is calling, and the scopes that remain after the chain of delegation. Scopes bound what the credential may do.
export interface Caller extends Principal {
  scopes: string[];
}

// The caller behind a request, from an agentic-identity `Auth`. The port reads the credential and resolves the chain,
// so a revoked or expired principal anywhere above the credential gives null.
export function fromAuth(auth: Auth): Authenticator {
  return async (request) => {
    const context = await auth.authenticate(request);
    if (context === null || !PRINCIPAL_KINDS.includes(context.kind) || !ID.test(context.principalId)) return null;
    return { id: context.principalId, kind: context.kind, scopes: context.scopes };
  };
}

const SEED_LABEL = "sheets-env";
const KEY_CHARS = /^[A-Za-z0-9_-]+$/;
const ROOT_KINDS: PrincipalKind[] = ["human", "agent"];

interface Seed {
  key: string;
  principal: IdentityPrincipal;
}

function seedsOf(keys: Record<string, { id?: unknown; kind?: unknown; parent?: unknown; scopes?: unknown }>): Seed[] {
  const named = new Map<string, PrincipalKind>();
  for (const entry of Object.values(keys)) {
    const who = principal(entry.id, entry.kind);
    if (who !== null) named.set(who.id, who.kind);
  }
  const seeds = Object.entries(keys).map(([key, entry]) => {
    if (!KEY_CHARS.test(key) || encoder.encode(key).length < MIN_KEY_BYTES) {
      throw new ConfigError(`an API key must have only letters, digits, "_" and "-", and at least ${MIN_KEY_BYTES} bytes`);
    }
    const who = principal(entry.id, entry.kind);
    if (who === null) throw new ConfigError(`each API key needs an id and a kind (${PRINCIPAL_KINDS.join(", ")})`);
    const scopes = entry.scopes === undefined ? ["*:admin"] : entry.scopes;
    if (!Array.isArray(scopes) || scopes.length === 0 || !scopes.every(isScope)) {
      throw new ConfigError('scopes must be a list like ["workbook:read"]');
    }
    const root = ROOT_KINDS.includes(who.kind);
    if (root !== (entry.parent === undefined)) {
      throw new ConfigError(root ? `a ${who.kind} has no parent` : `a ${who.kind} needs a "parent": the id of the principal it acts for`);
    }
    if (entry.parent !== undefined && (typeof entry.parent !== "string" || !named.has(entry.parent))) {
      throw new ConfigError(`the parent of "${who.id}" must be the id of another key's principal`);
    }
    return { key, principal: { id: who.id, kind: who.kind, scopes: scopes as string[], ...(entry.parent === undefined ? {} : { parentId: entry.parent as string }) } };
  });
  if (seeds.length === 0) throw new ConfigError("SHEETS_API_KEYS holds no keys");
  return seeds;
}

// Write the keys of the env into the store as principals and credentials, then drop each credential that an earlier
// env wrote and this one lacks. A credential doc is keyed by the hash of its key, so the store holds no key.
// Credentials that another tool wrote (no `sheets-env` label) stay.
export async function seedKeys(store: Store, keys: Record<string, { id?: unknown; kind?: unknown; parent?: unknown; scopes?: unknown }>): Promise<void> {
  const seeds = seedsOf(keys);
  const live = new Set<string>();
  for (const { key, principal: who } of seeds) {
    const hash = await digestHex(key);
    live.add(hash);
    try {
      await store.put("principals", who.id, { ...who });
    } catch {
      throw new ConfigError(`"${who.id}" is not a valid principal id for the store`);
    }
    await store.put("credentials", hash, { principalId: who.id, kind: "api_key", label: SEED_LABEL });
  }
  for (let cursor: string | undefined; ; ) {
    const page = await store.query("credentials", { where: { label: SEED_LABEL }, ...(cursor === undefined ? {} : { cursor }) });
    for (const doc of page.items) if (!live.has(doc.id)) await store.delete("credentials", doc.id);
    if (page.nextCursor === undefined || page.nextCursor === null) break;
    cursor = page.nextCursor;
  }
}

// Trust the caller's own claim. Anyone who can reach the Worker can write as anyone. Tests only.
export const trustHeaders: Authenticator = async (request) =>
  principal(request.headers.get("x-principal-id"), request.headers.get("x-principal-kind"));

export interface AuthEnv {
  /**
   * JSON: key to `{ "id", "kind", "parent"?, "scopes"? }`. The Worker writes these into the store as principals and
   * credentials, and reads every request through the `Auth` port of agentic-identity. A kind that acts for another
   * principal (`agent_of_human`, `agent_of_agent`, `user_token`) needs a `parent`: the id of another entry.
   */
  SHEETS_API_KEYS?: string;
  SHEETS_UNSAFE_TRUST_HEADERS?: string;
}

// The Worker has no login of its own, so it refuses to serve unless one of two things is set.
export async function authFromEnv(env: AuthEnv, store: Store): Promise<Authenticator> {
  const keys = env.SHEETS_API_KEYS;
  const unsafe = env.SHEETS_UNSAFE_TRUST_HEADERS === "1";
  if (keys !== undefined && keys !== "" && unsafe) {
    throw new ConfigError("set SHEETS_API_KEYS or SHEETS_UNSAFE_TRUST_HEADERS=1, not both");
  }
  if (keys !== undefined && keys !== "") {
    let parsed: unknown;
    try {
      parsed = JSON.parse(keys);
    } catch {
      throw new ConfigError('SHEETS_API_KEYS must be JSON: {"<key>": {"id": "...", "kind": "agent"}}');
    }
    if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
      throw new ConfigError('SHEETS_API_KEYS must be JSON: {"<key>": {"id": "...", "kind": "agent"}}');
    }
    await seedKeys(store, parsed as Record<string, { id?: unknown; kind?: unknown }>);
    return fromAuth(createStoreAuth(store));
  }
  if (unsafe) return trustHeaders;
  throw new ConfigError(
    "no auth is set. Set SHEETS_API_KEYS, or set SHEETS_UNSAFE_TRUST_HEADERS=1 to trust the x-principal-id and x-principal-kind headers (tests only)",
  );
}
