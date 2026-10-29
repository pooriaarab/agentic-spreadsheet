import { closeSync, existsSync, linkSync, utimesSync, mkdirSync, openSync, readdirSync, readFileSync, readSync, renameSync, rmSync, statSync, writeFileSync } from "node:fs";
import { createHash, randomUUID } from "node:crypto";
import { dirname, join } from "node:path";
import {
  assertBlobPutOptions,
  assertCollection,
  assertId,
  assertResponseHeaders,
  parseRegion,
  type AdapterOptions,
  type Blobs,
  type Doc,
  type Store,
  StoreConflictError,
} from "agentic-identity";
import { createFileStore as createIdentityFileStore } from "agentic-identity/file";

export type FileStoreOptions = AdapterOptions;

const isObject = (value: unknown): value is Record<string, unknown> => typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * What earlier versions left in `dir`, as documents by collection. Version 0.1.0 kept one file per workbook at
 * <dir>/workbooks/<id>.json. Later versions kept <dir>/store.json as `{ clock, tables: { name: { id: { v, doc } } } }`.
 * A file that is not a JSON object is an error, so a broken workbook is never dropped in the move.
 */
function readOld(dir: string): Map<string, Map<string, Doc>> | null {
  const path = join(dir, "store.json");
  const tables = new Map<string, Map<string, Doc>>();
  if (existsSync(path)) {
    let raw: unknown;
    try {
      raw = JSON.parse(readFileSync(path, "utf8"));
    } catch {
      throw new Error(`"${path}" is not an agentic-spreadsheet data file`);
    }
    if (isObject(raw) && raw.format !== undefined) return null;
    if (!isObject(raw) || !isObject(raw.tables)) throw new Error(`"${path}" is not an agentic-spreadsheet data file`);
    for (const [name, rows] of Object.entries(raw.tables)) {
      if (!isObject(rows)) throw new Error(`"${path}" is not an agentic-spreadsheet data file`);
      const table = new Map<string, Doc>();
      for (const [id, row] of Object.entries(rows)) {
        if (!isObject(row) || !isObject(row.doc)) throw new Error(`"${path}" is not an agentic-spreadsheet data file`);
        try {
          assertCollection(name);
          assertId(id);
        } catch {
          throw new Error(`"${path}" has a name or id that this version does not accept: "${name}/${id}"`);
        }
        table.set(id, row.doc as Doc);
      }
      tables.set(name, table);
    }
    return tables;
  }
  const folder = join(dir, "workbooks");
  if (!existsSync(folder)) return null;
  const rows = new Map<string, Doc>();
  for (const name of readdirSync(folder).filter((n) => n.endsWith(".json")).sort()) {
    const file = join(folder, name);
    let doc: unknown;
    try {
      doc = JSON.parse(readFileSync(file, "utf8"));
    } catch {
      throw new Error(`"${file}" is not valid JSON`);
    }
    if (!isObject(doc)) throw new Error(`"${file}" is not a workbook file`);
    const id = name.slice(0, -5);
    try {
      assertId(id);
    } catch {
      throw new Error(`"${file}" has an id that this version does not accept. Rename the file`);
    }
    rows.set(id, doc as Doc);
  }
  if (rows.size > 0) tables.set("workbooks", rows);
  return tables;
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
const MAX_ATTEMPTS = 50;

/** Runs `use` again while it loses to another writer, as the Store port asks of a transaction. Other errors pass. */
async function retrying<R>(use: () => Promise<R>): Promise<R> {
  for (let attempt = 1; ; attempt += 1) {
    try {
      return await use();
    } catch (error) {
      if (!(error instanceof StoreConflictError) || attempt >= MAX_ATTEMPTS) throw error;
      await sleep(Math.random() * Math.min(10 * attempt, 100));
    }
  }
}

/** True when no process has this id. Any other result, including no permission, counts as alive. */
function isDead(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return false;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "ESRCH";
  }
}

/**
 * Takes `<path>.migrating` for the length of `use`, so two processes do not move the same data. The lock holds its owner's
 * process id. One that is ten seconds old, or whose owner is gone, is moved aside by a rename, so one waiter wins.
 */
async function withMigrationLock<R>(path: string, use: () => Promise<R>): Promise<R> {
  const lock = `${path}.migrating`;
  const deadline = Date.now() + 30_000;
  mkdirSync(dirname(path), { recursive: true });
  for (;;) {
    try {
      writeFileSync(lock, `${process.pid}:`, { flag: "wx", mode: 0o600 });
      break;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    }
    try {
      const seen = readFileSync(lock, "utf8");
      const pid = Number(/^(\d+):/.exec(seen)?.[1]);
      if (Date.now() - statSync(lock).mtimeMs > 10_000 || (Number.isInteger(pid) && pid > 0 && isDead(pid))) {
        const aside = `${lock}.stale.${randomUUID()}`;
        renameSync(lock, aside);
        try {
          // A lock that changed in between is a newer one. Put it back unless another already exists.
          if (readFileSync(aside, "utf8") !== seen) {
            try {
              linkSync(aside, lock);
            } catch {}
          }
        } finally {
          rmSync(aside, { force: true });
        }
        continue;
      }
    } catch {
      continue; // The lock went away between the calls. Try to take it.
    }
    if (Date.now() > deadline) throw new StoreConflictError(`could not lock "${lock}" within 30000 ms`);
    await sleep(20 + Math.random() * 30);
  }
  // Touch the lock while the move runs, so a slow one is not taken over as stale.
  const beat = setInterval(() => {
    try {
      const now = new Date();
      utimesSync(lock, now, now);
    } catch {}
  }, 2_000);
  beat.unref();
  try {
    return await use();
  } finally {
    clearInterval(beat);
    rmSync(lock, { force: true });
  }
}

/**
 * A Store kept in `<dir>/store.json`, from agentic-identity. Data that an earlier version wrote there is moved into that format
 * on the first call. Stop processes that run an earlier version before this one starts. Calls on one store run one at a time,
 * so a call to the store inside a transaction callback waits for ever: use the `tx` it is given.
 */
export function createFileStore(dir: string, options: FileStoreOptions = {}): Store {
  const path = join(dir, "store.json");
  const store = createIdentityFileStore({ ...options, path });
  const migrate = () =>
    withMigrationLock(path, async () => {
      const old = readOld(dir); // Read again under the lock: another process may have moved it.
      if (old !== null) {
        // Build the new file beside the old one and rename it over, so a crash leaves one whole file.
        const next = `${path}.${randomUUID()}.moving`;
        try {
          const target = createIdentityFileStore({ ...options, path: next });
          await target.transaction(async (tx) => {
            for (const [collection, rows] of old) for (const [id, doc] of rows) await tx.put(collection, id, doc);
          });
          // An empty store still gets a file, so a leftover 0.1.0 folder is not read again.
          if (!existsSync(next)) await target.put("workbooks", "probe", {}).then(() => target.delete("workbooks", "probe"));
          renameSync(next, path);
        } finally {
          rmSync(next, { force: true });
        }
      }
    });
  /** True when old data may be here. A file in the new format starts with its `format` key, so a read of it needs no lock and writes nothing. */
  const mayNeedMove = (): boolean => {
    if (!existsSync(path)) return existsSync(join(dir, "workbooks"));
    const fd = openSync(path, "r");
    try {
      const head = Buffer.alloc(10);
      readSync(fd, head, 0, 10, 0);
      return head.toString("utf8") !== '{"format":';
    } finally {
      closeSync(fd);
    }
  };
  let moved: Promise<void> | undefined;
  const ready = () => {
    if (!mayNeedMove()) return Promise.resolve();
    moved ??= migrate().finally(() => {
      moved = undefined;
    });
    return moved;
  };
  return {
    region: store.region,
    get: async (collection, id) => (await ready(), retrying(() => store.get(collection, id))),
    query: async (collection, query) => (await ready(), retrying(() => store.query(collection, query))),
    put: async (collection, id, doc) => (await ready(), retrying(() => store.put(collection, id, doc))),
    delete: async (collection, id) => (await ready(), retrying(() => store.delete(collection, id))),
    transaction: async (fn) => (await ready(), retrying(() => store.transaction(fn))),
  };
}

/** Blobs kept as files under `dir`. Safe within one process only: the check and the write of a conditional put are not one step across processes, and the content type file is written after the data. File names are key hashes, so no key escapes the folder. */
export function createFileBlobs(dir: string, options: AdapterOptions = {}): Blobs {
  const path = (key: string) => join(dir, createHash("sha256").update(key).digest("hex"));
  const need = (key: string) => {
    if (key === "") throw new Error("blob key must not be empty");
  };
  const now = options.now ?? Date.now;
  // The version names the stored bytes.
  const versionOf = (data: Uint8Array) => createHash("sha256").update(data).digest("hex");
  // Another process may delete a file between a check and a read, so a missing file is the answer, not a failure.
  const attempt = <T>(use: () => T, none: T): T => {
    try {
      return use();
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return none;
      throw error;
    }
  };
  const typeOf = (key: string) => attempt(() => readFileSync(`${path(key)}.type`, "utf8") || undefined, undefined);
  const read = (key: string) => attempt(() => new Uint8Array(readFileSync(path(key))), null);
  return {
    region: options.region === undefined ? "unknown" : parseRegion(options.region),
    // Every step here is synchronous, so the check and the write cannot interleave with another caller in this process.
    async put(key, data, opts) {
      need(key);
      assertBlobPutOptions(opts);
      const current = read(key);
      if (opts?.ifNoneMatch === "*" && current !== null) return null;
      if (opts?.ifMatch !== undefined && (current === null || versionOf(current) !== opts.ifMatch)) return null;
      mkdirSync(dir, { recursive: true });
      const temp = `${path(key)}.${randomUUID()}.tmp`;
      writeFileSync(temp, data);
      renameSync(temp, path(key));
      if (opts?.contentType) writeFileSync(`${path(key)}.type`, opts.contentType);
      else rmSync(`${path(key)}.type`, { force: true });
      return { version: versionOf(data) };
    },
    async get(key) {
      need(key);
      const data = read(key);
      if (data === null) return null;
      const type = typeOf(key);
      return { data, version: versionOf(data), ...(type ? { contentType: type } : {}) };
    },
    async head(key) {
      need(key);
      const data = read(key);
      if (data === null) return null;
      const type = typeOf(key);
      return { size: data.length, version: versionOf(data), ...(type ? { contentType: type } : {}) };
    },
    async delete(key) {
      need(key);
      rmSync(path(key), { force: true });
      rmSync(`${path(key)}.type`, { force: true });
    },
    /** A local path with its expiry. Nothing serves or checks it: it exists so the port is whole. */
    async signedUrl(key, opts) {
      need(key);
      assertResponseHeaders(opts);
      if (!Number.isInteger(opts.expiresInSeconds) || opts.expiresInSeconds < 1 || opts.expiresInSeconds > 604800) {
        throw new RangeError("expiresInSeconds must be an integer from 1 to 604800");
      }
      const extra = [opts.responseContentDisposition, opts.responseContentType].map((v) => (v === undefined ? "" : encodeURIComponent(v)));
      return `file://${path(key)}?method=${opts.method}&expires=${now() + opts.expiresInSeconds * 1000}&d=${extra[0]}&t=${extra[1]}`;
    },
  };
}
