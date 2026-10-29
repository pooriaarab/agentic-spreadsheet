// The order of the rows, or of the columns, of one sheet.
//
// Every row has an id that never changes. A cell is keyed by the ids of its row and its column, so
// an insert, a delete, or a move touches no cell. A1 addresses are a view: the position of an id in
// the order.
//
// A sheet has about a million rows, so the order cannot list them all. The rows 1, 2, 3, ... that
// a sheet starts with are implicit: their ids are the numbers "1", "2", "3", and so on. Only an
// insert or a move writes an entry. An entry names the row it follows, so the order is a tree:
//
//   - A new row is a child of the row before it. Its siblings are listed newest first, so an insert
//     at a position lands before whatever sat at that position.
//   - Two writers who insert at the same place add two siblings. The merge lists them in one order
//     on both copies, and neither insert is lost.
//   - A row is deleted by marking its id gone. The mark is idempotent, so two deletes of one row agree.
//   - A move adds a new entry for the same id, in the new place. The latest entry for an id is the one
//     that counts. The old entries stay as anchors, so rows inserted after them keep their place.

export interface Entry {
  // The key of this entry. Other entries name it in `a`.
  k: string;
  // The id of the row. It equals `k` for the entry that made the row.
  id: string;
  // The key of the entry this one follows, or ROOT for the top of the sheet.
  a: string;
}

export const ROOT = "0";
const BASE_ID = /^[1-9]\d{0,9}$/;
const SAFE_ID = /^[0-9a-z.]{1,64}$/;

export const isBaseId = (id: string): boolean => BASE_ID.test(id);
export const isSafeId = (id: string): boolean => SAFE_ID.test(id);

function validEntry(value: unknown): value is Entry {
  const e = value as Partial<Entry> | null;
  return typeof e === "object" && e !== null && [e.k, e.id, e.a].every((part) => typeof part === "string" && SAFE_ID.test(part));
}

type Item = { run: true; from: number; to: number } | { run: false; id: string; key: string; visible: boolean };

// Where an id sits: how many visible rows come before its place, and whether the row itself shows.
export interface Place {
  before: number;
  visible: boolean;
}

export class Layout {
  private readonly items: Item[] = [];
  // For each item: the visible rows before it, and through it.
  private readonly starts: number[] = [];
  private readonly ends: number[] = [];
  // The item that holds each id, for the ids that are not part of a plain run.
  private readonly held = new Map<string, number>();
  // The item where each id was first placed, which a move does not change.
  private readonly origins = new Map<string, number>();
  private readonly runs: number[] = [];

  constructor(order: readonly unknown[], gone: ReadonlySet<string>) {
    // An entry key is used once, and is never a base id or the root. A repeat or a clash would
    // list the same children twice, and an entry that names itself would never end.
    const keys = new Set<string>();
    const entries = order.filter((value): value is Entry => validEntry(value) && !isBaseId(value.k) && value.k !== ROOT && !keys.has(value.k) && !!keys.add(value.k));
    const latest = new Map<string, Entry>();
    for (const entry of entries) latest.set(entry.id, entry);
    // Siblings, newest first: the entry made last comes first.
    const children = new Map<string, Entry[]>();
    for (let i = entries.length - 1; i >= 0; i -= 1) {
      const list = children.get(entries[i].a);
      if (list === undefined) children.set(entries[i].a, [entries[i]]);
      else list.push(entries[i]);
    }
    // A base row needs an item of its own when it has children, is gone, or has moved.
    const marked = new Set<number>();
    const mark = (id: string) => isBaseId(id) && marked.add(Number(id));
    for (const id of gone) mark(id);
    for (const entry of entries) (mark(entry.id), mark(entry.a));

    const emitChildren = (key: string) => {
      const stack: { list: Entry[]; at: number }[] = [{ list: children.get(key) ?? [], at: 0 }];
      while (stack.length > 0) {
        const top = stack[stack.length - 1];
        if (top.at >= top.list.length) {
          stack.pop();
          continue;
        }
        const entry = top.list[top.at++];
        const counts = latest.get(entry.id) === entry;
        this.push({ run: false, id: entry.id, key: entry.k, visible: counts && !gone.has(entry.id) }, counts);
        const below = children.get(entry.k);
        if (below !== undefined) stack.push({ list: below, at: 0 });
      }
    };
    emitChildren(ROOT);
    let previous = 0;
    for (const n of [...marked].sort((x, y) => x - y)) {
      if (n > previous + 1) this.push({ run: true, from: previous + 1, to: n - 1 }, false);
      const id = String(n);
      this.push({ run: false, id, key: id, visible: !gone.has(id) && !latest.has(id) }, !latest.has(id));
      emitChildren(id);
      previous = n;
    }
    this.push({ run: true, from: previous + 1, to: Infinity }, false);
  }

  private push(item: Item, holds: boolean): void {
    const at = this.items.length;
    const start = at === 0 ? 0 : this.ends[at - 1];
    this.items.push(item);
    this.starts.push(start);
    this.ends.push(start + (item.run ? item.to - item.from + 1 : item.visible ? 1 : 0));
    if (item.run) this.runs.push(at);
    else {
      if (holds) this.held.set(item.id, at);
      if (item.key === item.id) this.origins.set(item.id, at);
    }
  }

  // The first item whose rows reach position `p`.
  private find(p: number): number {
    let lo = 0;
    let hi = this.items.length - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.ends[mid] >= p) hi = mid;
      else lo = mid + 1;
    }
    return lo;
  }

  // The id of the row at 1-based position `p`.
  id(p: number): string {
    const at = this.find(p);
    const item = this.items[at];
    return item.run ? String(item.from + p - this.starts[at] - 1) : item.id;
  }

  // The key that new rows follow, to land right after the row at position `p`.
  key(p: number): string {
    const at = this.find(p);
    const item = this.items[at];
    return item.run ? String(item.from + p - this.starts[at] - 1) : item.key;
  }

  // The ids of `count` rows from position `from`.
  ids(from: number, count: number): string[] {
    const out: string[] = [];
    let p = from;
    for (let i = this.find(from); out.length < count; i += 1) {
      const item = this.items[i];
      if (item.run) {
        const offset = p - this.starts[i] - 1;
        const take = Math.min(count - out.length, item.to - item.from + 1 - offset);
        for (let n = 0; n < take; n += 1) out.push(String(item.from + offset + n));
        p += take;
      } else if (item.visible) {
        out.push(item.id);
        p += 1;
      }
    }
    return out;
  }

  // Where an id sits, or null for an id this order has never seen.
  place(id: string): Place | null {
    const at = this.held.get(id);
    if (at !== undefined) return { before: this.starts[at], visible: (this.items[at] as { visible: boolean }).visible };
    if (!isBaseId(id)) return null;
    const n = Number(id);
    for (let lo = 0, hi = this.runs.length - 1; lo <= hi; ) {
      const mid = (lo + hi) >> 1;
      const item = this.items[this.runs[mid]] as { from: number; to: number };
      if (n < item.from) hi = mid - 1;
      else if (n > item.to) lo = mid + 1;
      else return { before: this.starts[this.runs[mid]] + n - item.from, visible: true };
    }
    return null;
  }

  // Where an id was first placed. A range names its ends by this place, so a row that moves out of
  // a range leaves it, and a row that moves in joins it. A single ref uses `place` and follows the row.
  origin(id: string): Place | null {
    const at = this.origins.get(id);
    return at === undefined ? this.place(id) : { before: this.starts[at], visible: (this.items[at] as { visible: boolean }).visible };
  }

  // The 1-based position of a row that shows, or null.
  position(id: string): number | null {
    const place = this.place(id);
    return place === null || !place.visible ? null : place.before + 1;
  }
}
