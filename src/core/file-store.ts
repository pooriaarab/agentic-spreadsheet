import { link, mkdir, open, readdir, readFile, rename, rm, stat, utimes } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { join } from "node:path";
import { queryDocs, SAFE_NAME, type Doc, type QueryOptions, type QueryResult, type Store } from "./store.js";

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

export interface FileStoreOptions {
  // A lock whose file has not been renewed for this long is taken over, even if its process lives.
  staleMs?: number;
  // How long a transaction waits for the lock before it fails.
  waitMs?: number;
}

// One JSON file per document under <dir>/<collection>/<id>.json. A write goes to a temp file
// with its own name, then renames over the document, so a reader never sees half a file.
// A transaction holds <dir>/.lock, so processes and parallel calls take turns. The lock
// file names its process and is renewed while held. A lock from a dead process, or one that
// stops being renewed, is taken over.
export class FileStore implements Store {
  readonly jurisdiction = "local";
  private readonly staleMs: number;
  private readonly waitMs: number;

  constructor(
    private readonly dir: string,
    options: FileStoreOptions = {},
    private readonly held = false,
  ) {
    this.staleMs = options.staleMs ?? 30_000;
    this.waitMs = options.waitMs ?? 60_000;
  }

  private path(collection: string, id?: string): string {
    for (const part of [collection, id ?? "x"]) if (!SAFE_NAME.test(part)) throw new Error(`invalid store name "${part}"`);
    return id === undefined ? join(this.dir, collection) : join(this.dir, collection, `${id}.json`);
  }

  async get(collection: string, id: string): Promise<Doc | null> {
    try {
      return JSON.parse(await readFile(this.path(collection, id), "utf8")) as Doc;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
      throw error;
    }
  }

  async put(collection: string, id: string, doc: Doc): Promise<void> {
    const file = this.path(collection, id);
    await mkdir(this.path(collection), { recursive: true });
    const temp = `${file}.${process.pid}.${randomUUID()}.tmp`;
    const handle = await open(temp, "w");
    try {
      await handle.writeFile(JSON.stringify(doc));
      await handle.sync();
    } finally {
      await handle.close();
    }
    await rename(temp, file);
  }

  async delete(collection: string, id: string): Promise<void> {
    await rm(this.path(collection, id), { force: true });
  }

  async query(collection: string, options: QueryOptions = {}): Promise<QueryResult> {
    const names = await readdir(this.path(collection)).catch(() => [] as string[]);
    const ids = names.filter((n) => n.endsWith(".json")).map((n) => n.slice(0, -5));
    return queryDocs(ids, (id) => this.get(collection, id), options);
  }

  async transaction<T>(run: (tx: Store) => Promise<T>): Promise<T> {
    if (this.held) return run(this);
    const release = await this.lock();
    try {
      return await run(new FileStore(this.dir, { staleMs: this.staleMs, waitMs: this.waitMs }, true));
    } finally {
      await release();
    }
  }

  private async lock(): Promise<() => Promise<void>> {
    await mkdir(this.dir, { recursive: true });
    const file = join(this.dir, ".lock");
    const token = randomUUID();
    const deadline = Date.now() + this.waitMs;
    for (;;) {
      try {
        const handle = await open(file, "wx");
        await handle.writeFile(JSON.stringify({ pid: process.pid, token }));
        await handle.close();
        break;
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
      }
      await this.takeOverIfStale(file);
      if (Date.now() > deadline) throw new Error("the store is locked by another process");
      await sleep(5 + Math.random() * 20);
    }
    const renew = setInterval(() => void utimes(file, new Date(), new Date()).catch(() => undefined), Math.max(10, this.staleMs / 4));
    renew.unref();
    return async () => {
      clearInterval(renew);
      const owner = await readFile(file, "utf8").catch(() => "");
      if (owner.includes(token)) await rm(file, { force: true });
    };
  }

  // Move a stale lock aside with one atomic rename. Only one process wins the rename. If the file
  // we moved turns out to be a new lock, put it back with a link, which fails if another lock exists.
  private async takeOverIfStale(file: string): Promise<void> {
    let body: string;
    let age: number;
    try {
      [body, age] = [await readFile(file, "utf8"), Date.now() - (await stat(file)).mtimeMs];
    } catch {
      return;
    }
    let pid: number | undefined;
    try {
      pid = (JSON.parse(body) as { pid?: number }).pid;
    } catch {
      // A lock being written has no content yet. It is stale only when it has sat for the full period.
    }
    const dead = pid !== undefined && !alive(pid);
    if (!dead && age <= this.staleMs) return;
    const moved = `${file}.stale.${randomUUID()}`;
    try {
      await rename(file, moved);
    } catch {
      return;
    }
    if ((await readFile(moved, "utf8").catch(() => body)) !== body) await link(moved, file).catch(() => undefined);
    await rm(moved, { force: true });
  }
}

function alive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "EPERM";
  }
}
