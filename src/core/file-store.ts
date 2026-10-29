import { closeSync, existsSync, linkSync, mkdirSync, openSync, readdirSync, readFileSync, renameSync, rmSync, statSync, unlinkSync, writeFileSync } from "node:fs";
import { createHash, randomBytes, randomUUID } from "node:crypto";
import { join } from "node:path";
import {
  assertBlobPutOptions,
  assertCollection,
  assertResponseHeaders,
  assertId,
  createMemoryStore,
  parseRegion,
  type AdapterOptions,
  type Blobs,
  type Doc,
  type Page,
  type Query,
  type Store,
  StoreConflictError,
  type StoreOps,
} from "agentic-identity";

export interface FileStoreOptions extends AdapterOptions {
  /** A lock older than this is taken over, even if its owner looks alive. Default 10 seconds. */
  staleMs?: number;
  /** How long one operation waits for the lock before it fails. Default 30 seconds. */
  waitMs?: number;
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

/** True when no process has this id. Any other result, including no permission, counts as alive. */
function isDead(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return false;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "ESRCH";
  }
}

const MAX_TRANSACTION_ATTEMPTS = 50;

interface Row {
  /** Every write takes the next number, so a doc deleted and made again never repeats a version. */
  v: number;
  doc: Doc;
}
type Tables = Map<string, Map<string, Row>>;
interface Data {
  clock: number;
  tables: Tables;
}
interface Write {
  collection: string;
  id: string;
  doc: Doc | null;
}

const copy = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;

/**
 * A Store kept in one JSON file, safe for several processes. A write takes a
 * lock file, reads the file again, applies the change to that fresh copy, and
 * saves it with a write to a unique temp file and a rename. A read also reads
 * the file again, so a long-running process sees other writers.
 *
 * A transaction runs on a snapshot without the lock. At commit it takes the
 * lock, checks that every doc it read still has the version it saw, and then
 * writes. A doc that changed makes it run again, as the Store port requires.
 * It passes `runStoreContract`.
 */
export function createFileStore(dir: string, options: FileStoreOptions = {}): Store {
  const path = join(dir, "store.json");
  const region = options.region === undefined ? "unknown" : parseRegion(options.region);
  const staleMs = options.staleMs ?? 10_000;
  const waitMs = options.waitMs ?? 30_000;
  const lockPath = `${path}.lock`;
  const token = `${process.pid}:${randomBytes(8).toString("hex")}`;

  /**
   * Version 0.1.0 kept one file per workbook at <dir>/workbooks/<id>.json. Until the first write
   * makes store.json, those documents are read from there. A read changes nothing on disk. A file
   * that is not a JSON object is an error, so a broken workbook is never dropped in the move.
   */
  const legacy = (): Data => {
    const data: Data = { clock: 0, tables: new Map() };
    const folder = join(dir, "workbooks");
    if (!existsSync(folder)) return data;
    const rows = new Map<string, Row>();
    for (const name of readdirSync(folder).filter((n) => n.endsWith(".json")).sort()) {
      const file = join(folder, name);
      let doc: unknown;
      try {
        doc = JSON.parse(readFileSync(file, "utf8"));
      } catch {
        throw new Error(`"${file}" is not valid JSON`);
      }
      if (typeof doc !== "object" || doc === null || Array.isArray(doc)) throw new Error(`"${file}" is not a workbook file`);
      rows.set(name.slice(0, -5), { v: ++data.clock, doc: doc as Doc });
    }
    if (rows.size > 0) data.tables.set("workbooks", rows);
    return data;
  };

  /** Reads the file and checks its shape. A version that is not a whole number up to the clock would hide a lost update. */
  const load = (): Data => {
    if (!existsSync(path)) return legacy();
    const bad = () => new Error(`"${path}" is not an agentic-spreadsheet data file`);
    const isObject = (value: unknown): value is Record<string, unknown> => typeof value === "object" && value !== null && !Array.isArray(value);
    const raw: unknown = JSON.parse(readFileSync(path, "utf8"));
    if (!isObject(raw) || !Number.isSafeInteger(raw.clock) || (raw.clock as number) < 0 || !isObject(raw.tables)) throw bad();
    const clock = raw.clock as number;
    const tables: Tables = new Map();
    for (const [name, rows] of Object.entries(raw.tables)) {
      if (!isObject(rows)) throw bad();
      const table = new Map<string, Row>();
      for (const [id, row] of Object.entries(rows)) {
        if (!isObject(row) || !Number.isSafeInteger(row.v) || (row.v as number) < 1 || (row.v as number) > clock || !isObject(row.doc)) throw bad();
        table.set(id, { v: row.v as number, doc: row.doc as Doc });
      }
      tables.set(name, table);
    }
    return { clock, tables };
  };
  const save = (data: Data) => {
    const temp = `${path}.${process.pid}.${randomBytes(6).toString("hex")}.tmp`;
    const tables = Object.fromEntries([...data.tables].map(([name, rows]) => [name, Object.fromEntries(rows)]));
    writeFileSync(temp, JSON.stringify({ clock: data.clock, tables }), { mode: 0o600 });
    renameSync(temp, path);
  };
  const apply = (data: Data, { collection, id, doc }: Write) => {
    if (doc === null) {
      data.tables.get(collection)?.delete(id);
      return;
    }
    if (!data.tables.has(collection)) data.tables.set(collection, new Map());
    data.tables.get(collection)!.set(id, { v: ++data.clock, doc });
  };

  /** Judges a lock from its content and age. Content that cannot be read counts as live until it is old. */
  const isStale = (content: string, modified: number): boolean => {
    if (Date.now() - modified > staleMs) return true;
    const pid = Number(/^(\d+):/.exec(content)?.[1]);
    return Number.isInteger(pid) && pid > 0 && isDead(pid);
  };

  /** Moves a stale lock aside. The rename is atomic, so one process wins. A lock that changed in between is put back. */
  const takeOver = (seen: string) => {
    const aside = `${lockPath}.stale.${process.pid}.${randomBytes(6).toString("hex")}`;
    try {
      renameSync(lockPath, aside);
    } catch {
      return;
    }
    try {
      if (readFileSync(aside, "utf8") !== seen) {
        // This was a newer lock, not the stale one. Restore it unless another already exists.
        try {
          linkSync(aside, lockPath);
        } catch {}
      }
    } finally {
      try {
        unlinkSync(aside);
      } catch {}
    }
  };

  const acquire = async () => {
    mkdirSync(dir, { recursive: true });
    const deadline = Date.now() + waitMs;
    for (;;) {
      try {
        const fd = openSync(lockPath, "wx", 0o600);
        try {
          writeFileSync(fd, `${token}:${Date.now()}`);
        } finally {
          closeSync(fd);
        }
        return;
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
      }
      try {
        const content = readFileSync(lockPath, "utf8");
        if (isStale(content, statSync(lockPath).mtimeMs)) {
          takeOver(content);
          continue;
        }
      } catch {
        continue; // The lock went away between the two calls. Try to take it.
      }
      if (Date.now() > deadline) throw new StoreConflictError(`could not lock "${path}" within ${waitMs} ms`);
      await sleep(20 + Math.random() * 30);
    }
  };

  const release = () => {
    try {
      if (readFileSync(lockPath, "utf8").startsWith(token)) unlinkSync(lockPath);
    } catch {}
  };

  // One locked section at a time inside this process; the lock file covers other processes.
  let queue: Promise<unknown> = Promise.resolve();
  const exclusive = <R>(run: () => Promise<R>): Promise<R> => {
    const next = queue.then(run, run);
    queue = next.catch(() => undefined);
    return next;
  };

  /** Runs `use` on a fresh copy of the file, holding the lock. `use` returns false to leave the file unsaved. */
  const locked = <R>(use: (data: Data) => R | Promise<R>, saves: (result: R) => boolean = () => true): Promise<R> =>
    exclusive(async () => {
      await acquire();
      try {
        const data = load();
        const result = await use(data);
        if (saves(result)) save(data);
        return result;
      } finally {
        release();
      }
    });

  /** A private memory store filled from the file. The memory adapter checks names, cursors, and the read-before-write rule. */
  const snapshot = async (data: Data): Promise<Store> => {
    const store = createMemoryStore({ region });
    for (const [collection, rows] of data.tables) for (const [id, row] of rows) await store.put(collection, id, row.doc);
    return store;
  };

  async function transaction<R>(fn: (tx: StoreOps) => Promise<R>): Promise<R> {
    for (let attempt = 0; attempt < MAX_TRANSACTION_ATTEMPTS; attempt++) {
      const data = load();
      const seen = new Map<string, { collection: string; id: string; v: number }>();
      const writes: Write[] = [];
      const note = (collection: string, id: string) => {
        const key = `${collection}/${id}`;
        if (!seen.has(key)) seen.set(key, { collection, id, v: data.tables.get(collection)?.get(id)?.v ?? 0 });
      };
      const result = await (await snapshot(data)).transaction(async (tx) => {
        seen.clear();
        writes.length = 0;
        return fn({
          async get<T extends Doc>(collection: string, id: string) {
            const found = await tx.get<T>(collection, id);
            note(collection, id);
            return found;
          },
          async put(collection, id, doc) {
            await tx.put(collection, id, doc);
            writes.push({ collection, id, doc: copy(doc) });
          },
          async delete(collection, id) {
            await tx.delete(collection, id);
            writes.push({ collection, id, doc: null });
          },
          async query<T extends Doc>(collection: string, query?: Query): Promise<Page<T>> {
            const page = await tx.query<T>(collection, query);
            for (const item of page.items) note(collection, item.id);
            return page;
          },
        });
      });
      if (seen.size === 0 && writes.length === 0) return result;
      const unchanged = await locked(
        (fresh) => {
          for (const { collection, id, v } of seen.values()) if ((fresh.tables.get(collection)?.get(id)?.v ?? 0) !== v) return false;
          for (const write of writes) apply(fresh, write);
          return true;
        },
        (ok) => ok && writes.length > 0,
      );
      if (unchanged) return result;
    }
    throw new StoreConflictError(`transaction kept conflicting after ${MAX_TRANSACTION_ATTEMPTS} attempts`);
  }

  return {
    region,
    async get<T extends Doc>(collection: string, id: string) {
      assertCollection(collection);
      assertId(id);
      const row = load().tables.get(collection)?.get(id);
      return row ? copy(row.doc as T) : null;
    },
    async query<T extends Doc>(collection: string, query?: Query) {
      return (await snapshot(load())).query<T>(collection, query);
    },
    async put(collection, id, doc) {
      assertCollection(collection);
      assertId(id);
      const stored = copy(doc);
      await locked((data) => apply(data, { collection, id, doc: stored }));
    },
    async delete(collection, id) {
      assertCollection(collection);
      assertId(id);
      await locked((data) => apply(data, { collection, id, doc: null }));
    },
    transaction,
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
