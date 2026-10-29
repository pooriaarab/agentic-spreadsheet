import { createMemoryStore } from "agentic-identity";
import { describe, expect, test } from "vitest";
import { authFromEnv, ConfigError } from "./auth.js";

// Ways the seed of env keys can fail, listed before the code:
// - it overwrites a principal that another tool wrote (changes its rights, lifts its revocation)
// - it lifts the revocation of a principal that an earlier env wrote
// - an old env's key comes back when a Worker with an older env seeds the shared store

const KEY = "seed-key-0123456789-0123456789-abcdefgh";
const OLD = "old-key-0123456789-0123456789-abcdefghij";
const call = (auth: Awaited<ReturnType<typeof authFromEnv>>, key: string) => auth(new Request("https://x.test/", { headers: { authorization: `Bearer ${key}` } }));
const env = (keys: object) => ({ SHEETS_API_KEYS: JSON.stringify(keys) });

describe("seeding env keys", () => {
  test("a principal that another tool wrote is refused, and left as it was", async () => {
    const store = createMemoryStore();
    await store.put("principals", "ana", { id: "ana", kind: "human", scopes: ["workbook:read"], revokedAt: 1 });
    await expect(authFromEnv(env({ [KEY]: { id: "ana", kind: "human" } }), store)).rejects.toThrow(ConfigError);
    expect(await store.get("principals", "ana")).toMatchObject({ scopes: ["workbook:read"], revokedAt: 1 });
  });

  test("a revoked principal that an earlier env wrote stays revoked", async () => {
    const store = createMemoryStore();
    expect(await call(await authFromEnv(env({ [KEY]: { id: "ana", kind: "human" } }), store), KEY)).toMatchObject({ id: "ana" });
    await store.put("principals", "ana", { ...(await store.get("principals", "ana"))!, revokedAt: 1 });
    expect(await call(await authFromEnv(env({ [KEY]: { id: "ana", kind: "human" } }), store), KEY)).toBeNull();
  });

  test("a key the env dropped stays dead when an older env seeds the store again", async () => {
    const store = createMemoryStore();
    const old = env({ [KEY]: { id: "ana", kind: "human" }, [OLD]: { id: "bob", kind: "human" } });
    const current = await authFromEnv(env({ [KEY]: { id: "ana", kind: "human" } }), store);
    await authFromEnv(old, store);
    expect(await call(current, OLD)).toBeNull();
    expect(await call(current, KEY)).toMatchObject({ id: "ana" });
  });
});
