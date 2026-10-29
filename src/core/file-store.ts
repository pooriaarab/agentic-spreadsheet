import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { createHash, randomUUID } from "node:crypto";
import { join } from "node:path";
import {
  assertBlobPutOptions,
  assertId,
  assertResponseHeaders,
  parseRegion,
  type AdapterOptions,
  type Blobs,
  type Doc,
  type Store,
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
    const raw: unknown = JSON.parse(readFileSync(path, "utf8"));
    if (!isObject(raw) || !isObject(raw.tables) || raw.format !== undefined) return null;
    for (const [name, rows] of Object.entries(raw.tables)) {
      if (!isObject(rows)) throw new Error(`"${path}" is not an agentic-spreadsheet data file`);
      const table = new Map<string, Doc>();
      for (const [id, row] of Object.entries(rows)) {
        if (!isObject(row) || !isObject(row.doc)) throw new Error(`"${path}" is not an agentic-spreadsheet data file`);
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

/** A Store kept in `<dir>/store.json`, from agentic-identity. Data that an earlier version wrote there is moved into that format on first use. */
export function createFileStore(dir: string, options: FileStoreOptions = {}): Store {
  const path = join(dir, "store.json");
  const store = createIdentityFileStore({ ...options, path });
  let moved: Promise<void> | undefined;
  const migrate = async () => {
    const old = readOld(dir);
    if (old === null) return;
    // Build the new file beside the old one and rename it over, so a crash leaves one whole file.
    const next = `${path}.moving`;
    rmSync(next, { force: true });
    const target = createIdentityFileStore({ ...options, path: next });
    for (const [collection, rows] of old) for (const [id, doc] of rows) await target.put(collection, id, doc);
    if (existsSync(next)) renameSync(next, path);
    else rmSync(path, { force: true }); // An old file with no documents.
  };
  const ready = () => {
    if (!existsSync(path) && !existsSync(join(dir, "workbooks"))) return Promise.resolve();
    moved ??= migrate().catch((error: unknown) => {
      moved = undefined;
      throw error;
    });
    return moved;
  };
  return {
    region: store.region,
    get: async (collection, id) => (await ready(), store.get(collection, id)),
    query: async (collection, query) => (await ready(), store.query(collection, query)),
    put: async (collection, id, doc) => (await ready(), store.put(collection, id, doc)),
    delete: async (collection, id) => (await ready(), store.delete(collection, id)),
    transaction: async (fn) => (await ready(), store.transaction(fn)),
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
