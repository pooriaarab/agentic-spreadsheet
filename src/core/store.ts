export type Doc = Record<string, unknown>;
export interface QueryOptions { where?: Doc; orderBy?: string; limit?: number; cursor?: string }
export interface QueryResult { items: { id: string; doc: Doc }[]; cursor?: string }

// The storage port, small enough for a Firestore or D1 adapter: equality filters,
// one orderBy, limit, cursor paging, transactions. `jurisdiction` names where data lives.
// Read-modify-write code must run inside `transaction`.
export interface Store {
  readonly jurisdiction: string;
  get(collection: string, id: string): Promise<Doc | null>;
  put(collection: string, id: string, doc: Doc): Promise<void>;
  delete(collection: string, id: string): Promise<void>;
  query(collection: string, options?: QueryOptions): Promise<QueryResult>;
  transaction<T>(run: (tx: Store) => Promise<T>): Promise<T>;
}

export const SAFE_NAME = /^[\w-][\w.-]*$/;

// Filter, order, and page documents for any store. With no filter and no orderBy it pages by
// id and loads only the documents on the page, so a list does not read every document.
export async function queryDocs(
  allIds: string[],
  get: (id: string) => Promise<Doc | null>,
  { where = {}, orderBy, limit, cursor }: QueryOptions,
): Promise<QueryResult> {
  const ids = [...allIds].sort((a, b) => a.localeCompare(b));
  const load = async (page: string[]) =>
    (await Promise.all(page.map(async (id) => ({ id, doc: await get(id) })))).filter(
      (item): item is { id: string; doc: Doc } => item.doc !== null,
    );
  const finish = <R>(rows: R[], start: number) => {
    const end = limit === undefined ? rows.length : start + limit;
    return { page: rows.slice(start, end), more: end < rows.length };
  };
  if (Object.keys(where).length === 0 && orderBy === undefined) {
    if (cursor !== undefined && !SAFE_NAME.test(cursor)) throw new Error(`invalid cursor "${cursor}"`);
    const after = cursor === undefined ? 0 : ids.findIndex((id) => id.localeCompare(cursor) > 0);
    const { page, more } = finish(ids.map((id) => ({ id })), after < 0 ? ids.length : after);
    const items = await load(page.map((row) => row.id));
    return more ? { items, cursor: page[page.length - 1]?.id } : { items };
  }
  const key = (item: { id: string; doc: Doc }) => (orderBy === undefined ? item.id : String(item.doc[orderBy] ?? ""));
  const rows = (await load(ids))
    .filter((item) => Object.entries(where).every(([field, value]) => item.doc[field] === value))
    .sort((a, b) => key(a).localeCompare(key(b)) || a.id.localeCompare(b.id));
  const at = cursor === undefined ? -1 : rows.findIndex((item) => item.id === cursor);
  if (cursor !== undefined && at < 0) throw new Error(`unknown cursor "${cursor}"`);
  const { page, more } = finish(rows, at + 1);
  return more ? { items: page, cursor: page[page.length - 1]?.id } : { items: page };
}

// Documents in memory. A transaction runs after the one before it ends, so a read-modify-write
// cannot interleave with another. Data lives as long as the object, in one process. It is for
// tests and development.
export class MemoryStore implements Store {
  readonly jurisdiction = "memory";
  private readonly data = new Map<string, Map<string, string>>();
  private tail: Promise<unknown> = Promise.resolve();

  private name(part: string): string {
    if (!SAFE_NAME.test(part)) throw new Error(`invalid store name "${part}"`);
    return part;
  }

  async get(collection: string, id: string): Promise<Doc | null> {
    const text = this.data.get(this.name(collection))?.get(this.name(id));
    return text === undefined ? null : (JSON.parse(text) as Doc);
  }

  async put(collection: string, id: string, doc: Doc): Promise<void> {
    const docs = this.data.get(this.name(collection)) ?? new Map<string, string>();
    this.data.set(collection, docs.set(this.name(id), JSON.stringify(doc)));
  }

  async delete(collection: string, id: string): Promise<void> {
    this.data.get(this.name(collection))?.delete(this.name(id));
  }

  query(collection: string, options: QueryOptions = {}): Promise<QueryResult> {
    return queryDocs([...(this.data.get(this.name(collection))?.keys() ?? [])], (id) => this.get(collection, id), options);
  }

  transaction<T>(run: (tx: Store) => Promise<T>): Promise<T> {
    const turn = this.tail.then(() => run(Object.assign(Object.create(this) as MemoryStore, { transaction: (inner: (tx: Store) => Promise<unknown>) => inner(this) })));
    this.tail = turn.catch(() => undefined);
    return turn;
  }
}
