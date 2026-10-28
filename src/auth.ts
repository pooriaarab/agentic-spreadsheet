import { PRINCIPAL_KINDS, type Principal, type PrincipalKind } from "./core/workbook.js";

// Who is calling. It returns null when the request has no valid credential.
export type Authenticator = (request: Request) => Promise<Principal | null>;

export class ConfigError extends Error {}

const MIN_KEY_BYTES = 32;
const ID = /^[\w.@:-]{1,128}$/;
const encoder = new TextEncoder();

async function digest(text: string): Promise<Uint8Array> {
  return new Uint8Array(await crypto.subtle.digest("SHA-256", encoder.encode(text)));
}

// Compare two digests of the same length without stopping at the first difference.
function same(a: Uint8Array, b: Uint8Array): boolean {
  let diff = a.length ^ b.length;
  for (let i = 0; i < a.length; i += 1) diff |= a[i] ^ (b[i] ?? 0);
  return diff === 0;
}

function principal(id: unknown, kind: unknown): Principal | null {
  if (typeof id !== "string" || !ID.test(id) || !PRINCIPAL_KINDS.includes(kind as PrincipalKind)) return null;
  return { id, kind: kind as PrincipalKind };
}

// Each key maps to one principal. A key shorter than 32 bytes, or a missing kind, is a
// configuration error. A request is matched against every key, so timing does not show which one is close.
export async function bearerAuth(keys: Record<string, { id?: unknown; kind?: unknown }>): Promise<Authenticator> {
  const entries = await Promise.all(
    Object.entries(keys).map(async ([key, who]) => {
      const named = principal(who.id, who.kind);
      if (!/^\S+$/.test(key) || encoder.encode(key).length < MIN_KEY_BYTES) {
        throw new ConfigError(`an API key must have no spaces and at least ${MIN_KEY_BYTES} bytes`);
      }
      if (named === null) throw new ConfigError(`each API key needs an id and a kind (${PRINCIPAL_KINDS.join(", ")})`);
      return { hash: await digest(key), named };
    }),
  );
  if (entries.length === 0) throw new ConfigError("SHEETS_API_KEYS holds no keys");
  return async (request) => {
    const token = /^Bearer (\S+)$/.exec(request.headers.get("authorization") ?? "")?.[1];
    if (token === undefined) return null;
    const hash = await digest(token);
    let found: Principal | null = null;
    for (const entry of entries) if (same(hash, entry.hash)) found = entry.named;
    return found;
  };
}

// Trust the caller's own claim. Anyone who can reach the Worker can write as anyone. Tests only.
export const trustHeaders: Authenticator = async (request) =>
  principal(request.headers.get("x-principal-id"), request.headers.get("x-principal-kind"));

export interface AuthEnv {
  SHEETS_API_KEYS?: string;
  SHEETS_UNSAFE_TRUST_HEADERS?: string;
}

// The Worker has no login of its own, so it refuses to serve unless one of two things is set.
export async function authFromEnv(env: AuthEnv): Promise<Authenticator> {
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
    return bearerAuth(parsed as Record<string, { id?: unknown; kind?: unknown }>);
  }
  if (unsafe) return trustHeaders;
  throw new ConfigError(
    "no auth is set. Set SHEETS_API_KEYS, or set SHEETS_UNSAFE_TRUST_HEADERS=1 to trust the x-principal-id and x-principal-kind headers (tests only)",
  );
}
