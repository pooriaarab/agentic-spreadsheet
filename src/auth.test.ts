import { createMemoryStore } from "agentic-identity";
import { describe, expect, test } from "vitest";
import { authFromEnv, ConfigError } from "./auth.js";
import { revokeCredential } from "agentic-identity";

// Ways the seed of env keys can fail, listed before the code:
// - it overwrites a principal that another tool wrote (changes its rights, lifts its revocation)
// - it lifts the revocation of a principal that an earlier env wrote
// - it lifts the revocation of a credential, or a failed read of an existing doc is taken as "none"
// - a loop or a too-deep chain of parents starts and then never signs in
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

  test("a credential revoked in the store stays revoked after the Worker seeds again", async () => {
    const store = createMemoryStore();
    const keys = env({ [KEY]: { id: "ana", kind: "human" } });
    expect(await call(await authFromEnv(keys, store), KEY)).not.toBeNull();
    await revokeCredential(store, KEY, 1);
    expect(await call(await authFromEnv(keys, store), KEY)).toBeNull();
  });

  test("a failed read of an existing principal stops the start and writes nothing", async () => {
    const real = createMemoryStore();
    await real.put("principals", "ana", { id: "ana", kind: "human", scopes: ["workbook:read"], revokedAt: 1 });
    const broken = new Proxy(real, { get: (t, k) => (k === "get" ? async () => { throw new Error("d1 timeout"); } : Reflect.get(t, k)) });
    await expect(authFromEnv(env({ [KEY]: { id: "ana", kind: "human" } }), broken)).rejects.toThrow("d1 timeout");
    expect(await real.get("principals", "ana")).toMatchObject({ scopes: ["workbook:read"], revokedAt: 1 });
  });

  test("a loop of parents, a chain over the depth limit, and a key over 512 bytes are refused", async () => {
    const store = createMemoryStore();
    const k = (n: string) => `${n}-key-0123456789-0123456789-abcdefgh`;
    const loop = { [k("a")]: { id: "a", kind: "agent_of_agent", parent: "b" }, [k("b")]: { id: "b", kind: "agent_of_agent", parent: "a" } };
    await expect(authFromEnv(env(loop), store)).rejects.toThrow(/loop/);
    const deep: Record<string, object> = { [k("p0")]: { id: "p0", kind: "agent" } };
    for (let i = 1; i <= 5; i += 1) deep[k(`p${i}`)] = { id: `p${i}`, kind: "agent_of_agent", parent: `p${i - 1}` };
    await expect(authFromEnv(env(deep), store)).rejects.toThrow(/more than 4 parents/);
    await expect(authFromEnv(env({ ["x".repeat(513)]: { id: "a", kind: "agent" } }), store)).rejects.toThrow(/512 bytes/);
  });
});
