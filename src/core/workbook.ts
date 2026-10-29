import * as Y from "yjs";
import { PRINCIPAL_KINDS, type PrincipalKind } from "agentic-identity";
import { colIndex, colLetters, formatCell, MAX_COL, MAX_ROW, parseCell, parseRange, type RangePos } from "./a1.js";
import { Layout, ROOT, type Entry } from "./order.js";
import { hasControl, mapPlain, packRef, unpackRefs, type IdRef } from "./ref-ids.js";
import { formatSpan, mapRefs, quoteSheet, sameSheet, type Flags } from "./refs.js";

export { PRINCIPAL_KINDS, type PrincipalKind };

// The principal that makes an edit. The kind stays on every cell it writes.
export interface Principal {
  id: string;
  kind: PrincipalKind;
}

// Only a human edit is human-made. Agent edits, and edits a token
// makes on behalf of a user, are labeled as machine-made.
export function isMachineMade(kind: PrincipalKind): boolean {
  return kind !== "human";
}

export type CellValue = string | number | boolean;
// A string that starts with "=" is a formula. Null and "" clear the cell.
export type CellInput = CellValue | null;

export interface CellFormat {
  numberFormat?: string;
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center" | "right";
  color?: string;
  background?: string;
}

export type CellFormatPatch = { [K in keyof CellFormat]?: CellFormat[K] | null };

export interface Author extends Principal {
  machine: boolean;
  at: string;
}

export interface Cell {
  value: CellValue | null;
  formula: string | null;
  format: CellFormat;
  // Who last wrote the content (or cleared it). A format edit never changes it.
  author: Author | null;
  // Who last changed the format.
  formatAuthor: Author | null;
}

export interface SheetInfo {
  name: string;
  frozenRows: number;
  columnWidths: Record<string, number>;
}

export type Change = { all: true } | { all?: false; cells: { sheet: string; address: string }[] };

export class SheetError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SheetError";
  }
}

const MAX_RANGE_CELLS = 1_000_000;
const MAX_FORMAT_CELLS = 100_000;
export const MAX_STRUCTURE_COUNT = 10_000;
export const MAX_TEXT = 32_767;
export const MAX_SHEET_NAME = 256;
const MAX_FORMULA = 8_192;
const BAD_SHEET_NAME = /[[\]:*?/\\!']/;
const FORMAT_FIELDS = ["numberFormat", "bold", "italic", "align", "color", "background"] as const;
const COLOR = /^#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i;

type YSheet = Y.Map<unknown>;
// One flat map per sheet. A cell is a few keys: `<row id>|<col id>:c` (content), `:by` (the last
// content editor), `:fby` (the last format editor), and `:fmt.<field>`. Flat keys let two writers
// who touch one new cell merge. The ids never change, so a row or column edit moves no cell.
type Cells = Y.Map<unknown>;
type Stamp = Principal & { at: string };
type Axis = "row" | "col";

const ORDER_KEY = { row: "rowOrder", col: "colOrder" } as const;
const GONE_KEY = { row: "rowGone", col: "colGone" } as const;
const counters = new WeakMap<Y.Doc, number>();
// New ids, unique across every copy of the workbook: the client id of this copy, then a count.
function newIds(doc: Y.Doc, count: number): string[] {
  const first = counters.get(doc) ?? 0;
  counters.set(doc, first + count);
  return Array.from({ length: count }, (_, i) => `${doc.clientID.toString(36)}.${(first + i).toString(36)}`);
}

// A key in the old form has an A1 address before the colon, and no bar. Other keys are not cells.
const oldKey = (key: string): boolean => {
  if (key.includes("|")) return false;
  try {
    parseCell(key.slice(0, Math.max(0, key.indexOf(":"))));
    return true;
  } catch {
    return false;
  }
};

export class Workbook {
  readonly doc: Y.Doc;
  private readonly sheets: Y.Array<YSheet>;
  // The order of each sheet's rows and columns, and formulas shown as A1 text. A change to the
  // workbook clears both.
  private layouts = new Map<YSheet, Partial<Record<Axis, Layout>>>();
  private readonly shown = new Map<string, string>();
  // True when a sheet that has an id holds a cell key in the old A1 form. Such a cell reads as empty.
  // On purpose, the flag stays set until the workbook is loaded again.
  private oldKeys = false;

  constructor(doc: Y.Doc = new Y.Doc()) {
    this.doc = doc;
    this.sheets = doc.getArray("sheets");
    this.checkKeys();
    this.sheets.observeDeep((events) => {
      for (const event of events) this.watchKeys(event);
      // A write to a cell or a width changes neither the order nor the names, so it keeps both caches.
      if (events.some(({ path }) => !(path.length === 2 && (path[1] === "cells" || path[1] === "widths")))) this.invalidate();
    });
    this.migrate();
  }

  // True when two copies each converted an old workbook and then merged. Each copy made its own
  // order for the same sheets, and the merge keeps one of them, so row edits made after the
  // conversion can be lost, rows can read in the wrong place, and two sheets can share an id.
  // Every read and write throws until the workbook is loaded again from one copy.
  // A sheet that comes without an id after a conversion is tagged by its name, so one copy that
  // converts it is no clash, and two copies that convert the same sheet are.
  conversionClash(): boolean {
    const late = new Set<string>();
    let full = 0;
    for (const tag of this.doc.getMap("conversions").keys()) {
      if (!tag.startsWith("late:")) full += 1;
      else if (late.has(tag.slice(5, tag.lastIndexOf(":")))) return true;
      else late.add(tag.slice(5, tag.lastIndexOf(":")));
    }
    return full > 1;
  }

  // A sheet that has an id is in the new form. A cell key in the old form on it would read as empty,
  // so the workbook refuses every read and write instead.
  private checkKeys(sheets: Iterable<YSheet> = this.sheets): void {
    for (const sheet of sheets) {
      if (sheet.has("id") && [...this.cells(sheet).keys()].some(oldKey)) this.oldKeys = true;
    }
  }

  // The same check for a change that arrives after the load: a key in a cells map, a cells map that
  // is replaced, or a sheet that is added.
  private watchKeys(event: Y.YEvent<any>): void {
    const { path } = event;
    if (event instanceof Y.YArrayEvent && path.length === 0) {
      for (const item of event.changes.added) this.checkKeys(item.content.getContent().filter((v): v is YSheet => v instanceof Y.Map));
    } else if (event instanceof Y.YMapEvent && path.length === 1 && (event.keysChanged.has("cells") || event.keysChanged.has("id"))) {
      const sheet = this.sheets.get(path[0] as number);
      if (sheet !== undefined) this.checkKeys([sheet]);
    } else if (event instanceof Y.YMapEvent && path.length === 2 && path[1] === "cells" && this.sheets.get(path[0] as number)?.has("id")) {
      for (const key of event.keysChanged) if (event.target.has(key) && oldKey(key)) this.oldKeys = true;
    }
  }

  private usable(): void {
    if (this.oldKeys) {
      throw new SheetError("a sheet in this workbook holds a cell key in the old A1 form, so those cells would read as empty; restore the workbook from a copy saved before that key arrived");
    }
    if (this.sheets.toArray().some((sheet) => !sheet.has("id"))) {
      throw new SheetError("this workbook has a sheet in the old A1 form; load it with Workbook.fromState so it converts");
    }
    if (this.conversionClash()) {
      throw new SheetError("two copies converted an old workbook and then merged; load the workbook again from one copy");
    }
  }

  // Edits that run in one outer transaction fire no event until it ends, so each edit that changes
  // the order or the names clears the caches itself.
  private invalidate(): void {
    this.layouts = new Map();
    this.shown.clear();
  }

  // A workbook saved before cells had ids keys its cells by A1 address. Give each row and column
  // its number as its id, and key the cells by those ids. Each copy that converts makes its own
  // order for every sheet, and a merge keeps one of them. So the conversion tags the workbook with
  // the client id of the copy, and two tags mean a clash. Convert once, save, then share.
  private migrate(): void {
    const all = this.sheets.toArray();
    const legacy = all.filter((sheet) => !sheet.has("id"));
    if (legacy.length === 0) return;
    this.checkLegacyLog();
    this.convert(all, legacy);
  }

  // The old version kept a log of its row, column, and sheet-name edits. Two edits made at the same
  // time showed as a fork in it, and the old version refused writes until someone accepted the merge.
  // Cells in a fork may sit at the wrong address, and the conversion cannot tell, so it stops.
  private checkLegacyLog(): void {
    const log = this.doc.getArray<{ id?: unknown; parent?: unknown; accepted?: unknown } | null>("structure").toArray();
    let from = 0;
    log.forEach((entry, i) => entry?.accepted && (from = i + 1));
    for (let i = from; i < log.length; i += 1) {
      if (log[i]?.parent !== (i === 0 ? null : log[i - 1]?.id)) {
        throw new SheetError(
          "this old workbook had row, column, or sheet-name edits made at the same time and not accepted; its cells may sit at the wrong address; accept the merge in the version that saved it, then load it again",
        );
      }
    }
  }

  private convert(all: YSheet[], legacy: YSheet[]): void {
    // The stored form marks refs with U+0001 and U+0002. A formula that holds either would read as a
    // ref after the conversion. Check before any change, so a refusal leaves the workbook as it was.
    for (const sheet of legacy) {
      if (typeof sheet.get("name") !== "string") throw new SheetError("a sheet in the old A1 form has no name, so it cannot convert");
      for (const [key, value] of this.cells(sheet)) {
        if (key.endsWith(":c") && isFormula(value) && hasControl(value)) {
          throw new SheetError(`cell ${key.slice(0, -2)} on sheet "${sheet.get("name")}" holds a formula with a control character (U+0001 or U+0002), which the conversion cannot keep; remove the character in the version that saved it, then load it again`);
        }
      }
    }
    this.doc.transact(() => {
      // The client id is in each sheet id, so sheets from two conversions never share one. The colon
      // keeps a converted id apart from the ids that newIds makes, which use a dot.
      const prefix = this.doc.clientID.toString(36);
      // A workbook that was converted before has its tag. A sheet that came without an id after
      // that gets a tag with its own name, so the workbook does not read as a clash until a second
      // copy converts the same sheet.
      const tags = this.doc.getMap("conversions");
      if ([...tags.keys()].some((tag) => !tag.startsWith("late:"))) {
        for (const sheet of legacy) tags.set(`late:${(sheet.get("name") as string).toLowerCase()}:${prefix}`, true);
      } else tags.set(prefix, true);
      all.forEach((sheet, i) => sheet.has("id") || sheet.set("id", `${prefix}:s${i}`));
      const idOf = (name: string) => all.find((sheet) => sameSheet(sheet.get("name") as string, name))?.get("id") as string | undefined;
      for (const sheet of legacy) {
        for (const axis of ["row", "col"] as const) {
          if (!sheet.has(ORDER_KEY[axis])) sheet.set(ORDER_KEY[axis], new Y.Array());
          if (!sheet.has(GONE_KEY[axis])) sheet.set(GONE_KEY[axis], new Y.Map());
        }
        const home = sheet.get("name") as string;
        const cells = this.cells(sheet);
        for (const key of [...cells.keys()].sort()) {
          const cut = key.indexOf(":");
          let at: ReturnType<typeof parseCell>;
          try {
            at = parseCell(key.slice(0, cut));
          } catch {
            // A key that is not an address is dropped. One already in the new form is kept.
            if (!key.includes("|")) cells.delete(key);
            continue;
          }
          const value = cells.get(key);
          cells.delete(key);
          const carried = key.endsWith(":c") && isFormula(value) ? mapRefs(value, home, ({ text, ref, flags }) => {
            const sheetId = idOf(ref.target);
            const { c1, r1, c2, r2 } = ref.range;
            return sheetId === undefined ? text : packRef({ sheet: sheetId, r1: String(r1), c1: String(c1), r2: String(r2), c2: String(c2), qualified: ref.sheet !== null, flags });
          }) : value;
          cells.set(`${at.row}|${at.col}${key.slice(cut)}`, carried);
        }
        const widths = sheet.get("widths") as Y.Map<number>;
        for (const letters of [...widths.keys()].filter((key) => /^[A-Za-z]{1,3}$/.test(key)).sort()) {
          const width = widths.get(letters) as number;
          widths.delete(letters);
          widths.set(String(colIndex(letters)), width);
        }
        const frozen = sheet.get("frozenRows");
        if (typeof frozen === "number" && frozen > 0 && !sheet.has("frozen")) sheet.set("frozen", String(frozen));
        sheet.delete("frozenRows");
      }
      const log = this.doc.getArray("structure");
      log.delete(0, log.length);
    });
  }

  static fromState(state: Uint8Array): Workbook {
    const doc = new Y.Doc();
    Y.applyUpdate(doc, state);
    return new Workbook(doc);
  }

  encode(): Uint8Array {
    return Y.encodeStateAsUpdate(this.doc);
  }

  sheetNames(): string[] {
    this.usable();
    return this.sheets.toArray().map((sheet) => sheet.get("name") as string);
  }

  addSheet(name: string, index: number = this.sheets.length): void {
    this.usable();
    this.checkName(name);
    const sheet: YSheet = new Y.Map();
    sheet.set("id", newIds(this.doc, 1)[0]);
    sheet.set("name", name);
    sheet.set("cells", new Y.Map());
    sheet.set("widths", new Y.Map());
    for (const axis of ["row", "col"] as const) {
      sheet.set(ORDER_KEY[axis], new Y.Array());
      sheet.set(GONE_KEY[axis], new Y.Map());
    }
    this.sheets.insert(Math.max(0, Math.min(index, this.sheets.length)), [sheet]);
    this.invalidate();
    this.bindPending(sheet);
  }

  // A rename changes the name and nothing else. A formula names its sheet by id, so it follows.
  renameSheet(name: string, to: string): void {
    const sheet = this.find(name);
    const renamed = !sameSheet(name, to);
    if (renamed) this.checkName(to);
    sheet.set("name", to);
    this.invalidate();
    if (renamed) this.bindPending(sheet);
  }

  // A formula may name a sheet that does not exist yet. It stays as typed, and takes the sheet
  // once a sheet with that name appears.
  private bindPending(sheet: YSheet): void {
    const name = sheet.get("name") as string;
    const lower = name.toLowerCase();
    this.doc.transact(() => {
      for (const home of this.sheets) {
        const cells = this.cells(home);
        for (const [key, value] of [...cells.entries()]) {
          if (!key.endsWith(":c") || !isFormula(value) || !value.toLowerCase().includes(lower)) continue;
          const next = mapPlain(value, (plain) =>
            mapRefs(plain, home.get("name") as string, ({ text, ref, flags }) =>
              ref.sheet !== null && sameSheet(ref.target, name) ? this.pack(sheet, ref.range, ref.sheet !== null, flags) : text,
            ),
          );
          if (next !== value) cells.set(key, next);
        }
      }
    });
  }

  private pack(target: YSheet, range: RangePos, qualified: boolean, flags: Flags): string {
    const rows = this.layout(target, "row");
    const cols = this.layout(target, "col");
    const { c1, r1, c2, r2 } = range;
    // A range names places, a single ref names a row and a column.
    const name = flags.isRange ? (layout: Layout, at: number) => layout.key(at) : (layout: Layout, at: number) => layout.id(at);
    return packRef({ sheet: target.get("id") as string, r1: name(rows, r1), c1: name(cols, c1), r2: name(rows, r2), c2: name(cols, c2), qualified, flags });
  }

  // A formula that points at the deleted sheet reads as #REF!.
  deleteSheet(name: string): void {
    const sheet = this.find(name);
    if (this.sheets.length === 1) throw new SheetError("cannot delete the last sheet");
    this.sheets.delete(this.sheets.toArray().indexOf(sheet), 1);
    this.invalidate();
  }

  sheetInfo(name: string): SheetInfo {
    const sheet = this.find(name);
    const rows = this.layout(sheet, "row");
    const cols = this.layout(sheet, "col");
    const columnWidths: Record<string, number> = {};
    for (const [id, width] of (sheet.get("widths") as Y.Map<number>).entries()) {
      const at = cols.position(id);
      if (at !== null && at <= MAX_COL) columnWidths[colLetters(at)] = width;
    }
    const boundary = sheet.get("frozen");
    const place = typeof boundary === "string" ? rows.placeOf(boundary) : null;
    return {
      name: sheet.get("name") as string,
      frozenRows: place === null ? 0 : place.visible ? place.before + 1 : place.before,
      columnWidths,
    };
  }

  cellCount(name: string): number {
    return this.used(this.find(name)).length;
  }

  // The range from A1 to the last used cell, or null for an empty sheet.
  extent(name: string): string | null {
    const last = { col: 0, row: 0 };
    for (const at of this.used(this.find(name))) {
      last.col = Math.max(last.col, at.col);
      last.row = Math.max(last.row, at.row);
    }
    return last.col === 0 ? null : `A1:${formatCell(last)}`;
  }

  // Rows from the top stay in view. The sheet keeps the key of the place of the last frozen row, so
  // an insert above it, or a delete of a row in it, moves the line with the rows.
  setFrozenRows(name: string, rows: number): void {
    if (!Number.isInteger(rows) || rows < 0 || rows > MAX_ROW) throw new SheetError(`invalid frozen rows ${rows}`);
    const sheet = this.find(name);
    if (rows === 0) sheet.delete("frozen");
    else sheet.set("frozen", this.layout(sheet, "row").key(rows));
    this.invalidate();
  }

  setColumnWidth(name: string, column: string, width: number): void {
    const letters = column.toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(letters) || colIndex(letters) > MAX_COL) throw new SheetError(`invalid column "${column}"`);
    if (typeof width !== "number" || !Number.isFinite(width) || width <= 0) throw new SheetError(`invalid column width ${width}`);
    const sheet = this.find(name);
    (sheet.get("widths") as Y.Map<number>).set(this.layout(sheet, "col").id(colIndex(letters)), width);
  }

  readRange(name: string, range: string): Cell[][] {
    const sheet = this.find(name);
    const cells = this.cells(sheet);
    const span = this.span(range);
    const rows = this.layout(sheet, "row").ids(span.r1, span.r2 - span.r1 + 1);
    const cols = this.layout(sheet, "col").ids(span.c1, span.c2 - span.c1 + 1);
    return rows.map((row) => cols.map((col) => toCell(cells, `${row}|${col}`, (formula) => this.show(formula))));
  }

  // Tell a listener which cells changed content, or that anything may have changed. A change to
  // a format, a width, or an editor stamp does not count. The formula engine uses this to
  // recompute only what a change reaches, whether it came from here or from another copy.
  onChange(listener: (change: Change) => void): () => void {
    const handler = (events: Y.YEvent<any>[]) => {
      const cells: { sheet: string; address: string }[] = [];
      for (const event of events) {
        const { path } = event;
        const keys = event instanceof Y.YMapEvent ? [...event.keysChanged] : [];
        if (path.length === 2 && path[1] === "cells") {
          const sheet = this.sheets.get(path[0] as number);
          if (sheet === undefined) return listener({ all: true });
          const rows = this.layout(sheet, "row");
          const cols = this.layout(sheet, "col");
          for (const key of keys) {
            if (!key.endsWith(":c")) continue;
            const [row, col] = key.slice(0, -2).split("|");
            const at = { row: rows.position(row), col: cols.position(col) };
            if (at.row !== null && at.col !== null && at.row <= MAX_ROW && at.col <= MAX_COL) cells.push({ sheet: sheet.get("name") as string, address: formatCell({ col: at.col, row: at.row }) });
          }
        } else if (path.length === 1 ? keys.some((key) => key !== "frozen") : path[1] !== "widths") {
          return listener({ all: true });
        }
      }
      if (cells.length > 0) listener({ cells });
    };
    this.sheets.observeDeep(handler);
    return () => this.sheets.unobserveDeep(handler);
  }

  // A cell's content and nothing else. The formula engine reads cells through this.
  raw(name: string, col: number, row: number): { value: CellValue | null; formula: string | null } {
    const sheet = this.find(name);
    const content = this.cells(sheet).get(`${this.layout(sheet, "row").id(row)}|${this.layout(sheet, "col").id(col)}:c`);
    const valid = typeof content === "string" || typeof content === "boolean" || (typeof content === "number" && Number.isFinite(content));
    const formula = valid && isFormula(content) ? this.show(content) : null;
    return { value: valid && formula === null ? (content as CellValue) : null, formula };
  }

  // Write a grid of values, or formulas, with its top-left corner at `anchor`.
  writeRange(name: string, anchor: string, rows: CellInput[][], by: Principal): void {
    const sheet = this.find(name);
    const cells = this.cells(sheet);
    const start = parseCell(anchor);
    let width = 0;
    for (const line of rows) {
      width = Math.max(width, line.length);
      for (const value of line) checkValue(value);
    }
    if (start.row + rows.length - 1 > MAX_ROW || start.col + width - 1 > MAX_COL) {
      throw new SheetError(`write at ${anchor} runs past the grid`);
    }
    checkPrincipal(by);
    const rowIds = this.layout(sheet, "row").ids(start.row, rows.length);
    const colIds = this.layout(sheet, "col").ids(start.col, width);
    this.doc.transact(() => {
      rows.forEach((line, dr) =>
        line.forEach((value, dc) =>
          putValue(cells, `${rowIds[dr]}|${colIds[dc]}`, typeof value === "string" && isFormula(value) ? this.store(value, sheet) : value, by),
        ),
      );
    });
  }

  setFormat(name: string, range: string, patch: CellFormatPatch, by: Principal): void {
    const sheet = this.find(name);
    const cells = this.cells(sheet);
    const span = this.span(range, MAX_FORMAT_CELLS);
    checkPatch(patch);
    checkPrincipal(by);
    const rows = this.layout(sheet, "row").ids(span.r1, span.r2 - span.r1 + 1);
    const cols = this.layout(sheet, "col").ids(span.c1, span.c2 - span.c1 + 1);
    this.doc.transact(() => {
      for (const row of rows) {
        for (const col of cols) {
          const addr = `${row}|${col}`;
          for (const [field, value] of Object.entries(patch)) {
            if (value === null) cells.delete(`${addr}:fmt.${field}`);
            else if (value !== undefined) cells.set(`${addr}:fmt.${field}`, value);
          }
          tidyFormat(cells, addr, by);
        }
      }
    });
  }

  // Structure. `at` is the 1-based row or column. Each edit changes the order of the ids and
  // nothing else, so edits made at the same time all apply. Refs to moved cells follow them. Refs
  // to deleted cells read as #REF!. An insert that would push a cell or a width off the grid is refused.

  insertRows(name: string, at: number, count = 1): void {
    this.restructure(name, "row", "insert", at, count);
  }

  deleteRows(name: string, at: number, count = 1): void {
    this.restructure(name, "row", "delete", at, count);
  }

  insertCols(name: string, at: number, count = 1): void {
    this.restructure(name, "col", "insert", at, count);
  }

  deleteCols(name: string, at: number, count = 1): void {
    this.restructure(name, "col", "delete", at, count);
  }

  // Move `count` rows from `from` so they sit before the row now at `to`. Their cells, and refs to
  // them, go with them.
  moveRows(name: string, from: number, count: number, to: number): void {
    this.move(name, "row", from, count, to);
  }

  moveCols(name: string, from: number, count: number, to: number): void {
    this.move(name, "col", from, count, to);
  }

  private check(axis: Axis, at: number, count: number): number {
    const max = axis === "row" ? MAX_ROW : MAX_COL;
    if (!Number.isInteger(at) || !Number.isInteger(count) || at < 1 || at > max || count < 1 || count > MAX_STRUCTURE_COUNT) {
      throw new SheetError(`invalid ${axis} position ${at} or count ${count}`);
    }
    return max;
  }

  private restructure(name: string, axis: Axis, op: "insert" | "delete", at: number, count: number): void {
    const max = this.check(axis, at, count);
    // A delete must name rows or columns that exist, so a range past the edge is an error, not a no-op.
    if (op === "delete" && at + count - 1 > max) throw new SheetError(`cannot delete ${count} ${axis}s from ${at}: the sheet ends at ${max}`);
    const sheet = this.find(name);
    const layout = this.layout(sheet, axis);
    const order = sheet.get(ORDER_KEY[axis]) as Y.Array<Entry>;
    const widths = sheet.get("widths") as Y.Map<number>;
    if (op === "insert") {
      let last = 0;
      for (const used of this.used(sheet)) last = Math.max(last, used[axis]);
      if (axis === "col") for (const id of widths.keys()) last = Math.max(last, layout.position(id) ?? 0);
      if (last >= at && last + count > max) throw new SheetError(`insert would push cells past the edge of the grid`);
      const anchor = at === 1 ? ROOT : layout.key(at - 1);
      // Siblings read newest first, so the rows go in last to first.
      order.push(newIds(this.doc, count).reverse().map((id) => ({ k: id, id, a: anchor })));
      this.invalidate();
      return;
    }
    const dead = new Set(layout.ids(at, count));
    const cells = this.cells(sheet);
    const gone = sheet.get(GONE_KEY[axis]) as Y.Map<true>;
    this.doc.transact(() => {
      for (const id of dead) gone.set(id, true);
      for (const key of [...cells.keys()]) {
        const [row, col] = key.slice(0, key.indexOf(":")).split("|");
        if (dead.has(axis === "row" ? row : col)) cells.delete(key);
      }
      if (axis === "col") for (const id of dead) widths.delete(id);
    });
    this.invalidate();
  }

  private move(name: string, axis: Axis, from: number, count: number, to: number): void {
    const max = this.check(axis, from, count);
    if (!Number.isInteger(to) || to < 1 || to > max + 1) throw new SheetError(`invalid ${axis} position ${to}`);
    if (to >= from && to <= from + count) return;
    const sheet = this.find(name);
    const layout = this.layout(sheet, axis);
    const ids = layout.ids(from, count);
    const keys = newIds(this.doc, count);
    let after = to === 1 ? ROOT : layout.key(to - 1);
    const entries = ids.map((id, i) => {
      const entry = { k: keys[i], id, a: after };
      after = keys[i];
      return entry;
    });
    (sheet.get(ORDER_KEY[axis]) as Y.Array<Entry>).push(entries);
    this.invalidate();
  }

  private layout(sheet: YSheet, axis: Axis): Layout {
    let built = this.layouts.get(sheet);
    if (built === undefined) this.layouts.set(sheet, (built = {}));
    const order = sheet.get(ORDER_KEY[axis]) as Y.Array<unknown> | undefined;
    const gone = sheet.get(GONE_KEY[axis]) as Y.Map<unknown> | undefined;
    return (built[axis] ??= new Layout(order?.toArray() ?? [], new Set(gone?.keys() ?? [])));
  }

  // The positions of the cells that hold anything and sit on a row and a column that show, inside
  // the grid. Two inserts made at once can push a cell past the last row. It stays in the workbook,
  // unread, until rows go.
  private used(sheet: YSheet): { row: number; col: number }[] {
    const rows = this.layout(sheet, "row");
    const cols = this.layout(sheet, "col");
    const seen = new Set<string>();
    const out: { row: number; col: number }[] = [];
    for (const key of this.cells(sheet).keys()) {
      const base = key.slice(0, key.indexOf(":"));
      if (seen.has(base)) continue;
      seen.add(base);
      const [rowId, colId] = base.split("|");
      const row = rows.position(rowId);
      const col = cols.position(colId);
      if (row !== null && col !== null && row <= MAX_ROW && col <= MAX_COL) out.push({ row, col });
    }
    return out;
  }

  // Store a formula typed as A1 text with each ref turned into ids. A ref to a sheet that does not exist stays as typed.
  private store(formula: string, home: YSheet): string {
    return mapRefs(formula, home.get("name") as string, ({ text, ref, flags }) => {
      const target = ref.sheet === null ? home : this.sheets.toArray().find((s) => sameSheet(s.get("name") as string, ref.target));
      return target === undefined ? text : this.pack(target, ref.range, ref.sheet !== null, flags);
    });
  }

  // A stored formula as A1 text, for the rows and columns as they stand now.
  private show(stored: string): string {
    if (!stored.includes("\u0001")) return stored;
    let text = this.shown.get(stored);
    if (text === undefined) this.shown.set(stored, (text = unpackRefs(stored, (ref) => this.showRef(ref))));
    return text;
  }

  private showRef(ref: IdRef): string {
    const sheet = this.sheets.toArray().find((s) => s.get("id") === ref.sheet);
    if (sheet === undefined) return "#REF!";
    const rows = reach(this.layout(sheet, "row"), ref.r1, ref.r2, MAX_ROW, ref.flags.isRange);
    const cols = reach(this.layout(sheet, "col"), ref.c1, ref.c2, MAX_COL, ref.flags.isRange);
    if (rows === null || cols === null) return "#REF!";
    const prefix = ref.qualified ? `${quoteSheet(sheet.get("name") as string)}!` : "";
    return prefix + formatSpan({ c1: cols[0], c2: cols[1], r1: rows[0], r2: rows[1] }, ref.flags);
  }

  private find(name: string): YSheet {
    this.usable();
    const sheet = this.sheets.toArray().find((s) => sameSheet(s.get("name") as string, name));
    if (sheet === undefined) throw new SheetError(`no sheet "${name}"`);
    return sheet;
  }

  private cells(sheet: YSheet): Cells {
    return sheet.get("cells") as Cells;
  }

  private checkName(name: string): void {
    if (name.trim() === "" || name.length > MAX_SHEET_NAME || BAD_SHEET_NAME.test(name)) throw new SheetError(`invalid sheet name "${name}"`);
    if (this.sheetNames().some((existing) => sameSheet(existing, name))) {
      throw new SheetError(`sheet "${name}" exists`);
    }
  }

  private span(range: string, max = MAX_RANGE_CELLS): RangePos {
    const span = parseRange(range);
    if ((span.c2 - span.c1 + 1) * (span.r2 - span.r1 + 1) > max) {
      throw new SheetError(`range ${range} is too large`);
    }
    return span;
  }
}

// The positions a ref spans on one axis. A range names each end by a place. An end whose row is
// deleted or moved away leaves the range, which then stops at the nearest row that shows. A span with no row left, or one that starts past the edge of the grid, has none. A span
// that reaches past the edge is cut there.
function reach(layout: Layout, from: string, to: string, max: number, range: boolean): [number, number] | null {
  const a = range ? layout.placeOf(from) : layout.place(from);
  const b = range ? layout.placeOf(to) : layout.place(to);
  if (a === null || b === null) return null;
  const lo = a.before + 1;
  const hi = b.visible ? b.before + 1 : b.before;
  return lo > max || hi < lo ? null : [lo, Math.min(hi, max)];
}

function checkPrincipal(by: Principal): void {
  if (by.id === "" || !PRINCIPAL_KINDS.includes(by.kind)) throw new SheetError("invalid principal");
}

const isFormula = (v: unknown): v is string => typeof v === "string" && v.startsWith("=") && v.length > 1;

function checkValue(value: unknown): void {
  if (value === null || typeof value === "boolean") return;
  if (typeof value === "number" && Number.isFinite(value)) return;
  if (typeof value !== "string") throw new SheetError("invalid cell value");
  if (isFormula(value) && value.length > MAX_FORMULA) throw new SheetError(`formula longer than ${MAX_FORMULA} characters`);
  if (isFormula(value) && hasControl(value)) throw new SheetError("a formula may not hold a control character");
  if (value.length > MAX_TEXT) throw new SheetError(`text longer than ${MAX_TEXT} characters`);
}

function checkPatch(patch: CellFormatPatch): void {
  for (const [field, value] of Object.entries(patch)) {
    if (!(FORMAT_FIELDS as readonly string[]).includes(field)) throw new SheetError(`unknown format field "${field}"`);
    if (value === null || value === undefined) continue;
    const ok =
      field === "numberFormat" ? typeof value === "string" && value.length > 0 && value.length <= 100
      : field === "align" ? value === "left" || value === "center" || value === "right"
      : field === "bold" || field === "italic" ? typeof value === "boolean"
      : typeof value === "string" && COLOR.test(value);
    if (!ok) throw new SheetError(`invalid ${field === "color" || field === "background" ? `${field} color` : field}`);
  }
}

function stamp(cells: Cells, key: string, by: Principal): void {
  cells.set(key, { id: by.id, kind: by.kind, at: new Date().toISOString() });
}

const hasFormat = (cells: Cells, addr: string) => FORMAT_FIELDS.some((field) => cells.has(`${addr}:fmt.${field}`));

// After a format edit: stamp the format editor while a format remains. Leave the content editor alone.
function tidyFormat(cells: Cells, addr: string, by: Principal): void {
  if (hasFormat(cells, addr)) stamp(cells, `${addr}:fby`, by);
  else {
    cells.delete(`${addr}:fby`);
    if (!cells.has(`${addr}:c`)) cells.delete(`${addr}:by`);
  }
}

function putValue(cells: Cells, addr: string, value: CellInput, by: Principal): void {
  if (value === null || value === "") {
    if (!cells.has(`${addr}:c`)) return;
    cells.delete(`${addr}:c`);
    if (hasFormat(cells, addr)) stamp(cells, `${addr}:by`, by);
    else cells.delete(`${addr}:by`);
    return;
  }
  cells.set(`${addr}:c`, value);
  stamp(cells, `${addr}:by`, by);
}

const toAuthor = (by: Stamp | undefined): Author | null =>
  by === undefined ? null : { id: by.id, kind: by.kind, machine: isMachineMade(by.kind), at: by.at };

function toCell(cells: Cells, addr: string, show: (formula: string) => string): Cell {
  const by = cells.get(`${addr}:by`) as Stamp | undefined;
  const fby = cells.get(`${addr}:fby`) as Stamp | undefined;
  if (by === undefined && fby === undefined) return { value: null, formula: null, format: {}, author: null, formatAuthor: null };
  const content = cells.get(`${addr}:c`);
  const valid = typeof content === "string" || typeof content === "boolean" || (typeof content === "number" && Number.isFinite(content));
  const format: Record<string, unknown> = {};
  for (const field of FORMAT_FIELDS) {
    const value = cells.get(`${addr}:fmt.${field}`);
    if (value !== undefined) format[field] = value;
  }
  const formula = valid && isFormula(content) ? show(content) : null;
  return {
    value: valid && formula === null ? (content as CellValue) : null,
    formula,
    format: format as CellFormat,
    author: toAuthor(by),
    formatAuthor: toAuthor(fby),
  };
}
