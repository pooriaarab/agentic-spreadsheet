import * as Y from "yjs";
import { colIndex, colLetters, formatCell, MAX_COL, MAX_ROW, parseCell, parseRange, type RangePos } from "./a1.js";
import { rewriteRefs, sameSheet, shiftSpan, type RefEdit } from "./refs.js";

export type PrincipalKind = "human" | "agent_of_human" | "agent_of_agent" | "agent" | "user_token";

export const PRINCIPAL_KINDS: readonly PrincipalKind[] = ["human", "agent_of_human", "agent_of_agent", "agent", "user_token"];

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
const MAX_FORMULA = 8_192;
const BAD_SHEET_NAME = /[[\]:*?/\\!']/;
const FORMAT_FIELDS = ["numberFormat", "bold", "italic", "align", "color", "background"] as const;
const COLOR = /^#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i;

type YSheet = Y.Map<unknown>;
// One flat map per sheet. A cell is a few keys: `A1:c` (content), `A1:by` (the last
// content editor), `A1:fby` (the last format editor), and `A1:fmt.<field>`. Flat keys let two writers who touch one new cell merge.
type Cells = Y.Map<unknown>;
type Stamp = Principal & { at: string };
type Axis = "row" | "col";

// A structure edit (insert, delete, rename) moves cells and rewrites formulas, so it cannot
// merge with another one made at the same time. Each edit adds a log entry that names the
// entry before it. Two edits made at once name the same parent, and the merge shows the fork.
interface LogEntry {
  id: string;
  parent: string | null;
  accepted?: true;
}

export class Workbook {
  readonly doc: Y.Doc;
  private readonly sheets: Y.Array<YSheet>;
  private readonly log: Y.Array<LogEntry>;

  constructor(doc: Y.Doc = new Y.Doc()) {
    this.doc = doc;
    this.sheets = doc.getArray("sheets");
    this.log = doc.getArray("structure");
  }

  // True when two clients changed rows, columns, or sheet names at the same time. Their edits
  // cannot both hold, so the workbook refuses writes until someone checks it and accepts it.
  structureConflict(): boolean {
    const entries = this.log.toArray();
    let from = 0;
    entries.forEach((entry, i) => entry.accepted && (from = i + 1));
    for (let i = from; i < entries.length; i += 1) {
      if (entries[i].parent !== (i === 0 ? null : entries[i - 1].id)) return true;
    }
    return false;
  }

  acceptStructureMerge(): void {
    this.record(true);
  }

  private record(accepted?: true): void {
    const last = this.log.get(this.log.length - 1);
    const id = `${this.doc.clientID.toString(36)}.${Date.now().toString(36)}.${Math.random().toString(36).slice(2)}`;
    this.log.push([{ id, parent: last?.id ?? null, ...(accepted && { accepted }) }]);
  }

  private writable(): void {
    if (this.structureConflict()) {
      throw new SheetError("two clients changed rows, columns, or sheet names at the same time; check the sheet, then accept the merge");
    }
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
    return this.sheets.toArray().map((sheet) => sheet.get("name") as string);
  }

  addSheet(name: string, index: number = this.sheets.length): void {
    this.writable();
    this.checkName(name);
    const sheet: YSheet = new Y.Map();
    sheet.set("name", name);
    sheet.set("cells", new Y.Map());
    sheet.set("widths", new Y.Map());
    sheet.set("frozenRows", 0);
    this.sheets.insert(Math.max(0, Math.min(index, this.sheets.length)), [sheet]);
  }

  renameSheet(name: string, to: string): void {
    this.writable();
    const sheet = this.find(name);
    if (!sameSheet(name, to)) this.checkName(to);
    this.doc.transact(() => {
      this.rewriteAll((ref) =>
        ref.sheet !== null && sameSheet(ref.target, name) ? { range: ref.range, sheet: to } : { range: ref.range },
      );
      sheet.set("name", to);
      this.record();
    });
  }

  deleteSheet(name: string): void {
    this.writable();
    const sheet = this.find(name);
    if (this.sheets.length === 1) throw new SheetError("cannot delete the last sheet");
    this.doc.transact(() => {
      this.rewriteAll((ref) => (sameSheet(ref.target, name) ? null : { range: ref.range }));
      this.sheets.delete(this.sheets.toArray().indexOf(sheet), 1);
      this.record();
    });
  }

  sheetInfo(name: string): SheetInfo {
    const sheet = this.find(name);
    return {
      name: sheet.get("name") as string,
      frozenRows: sheet.get("frozenRows") as number,
      columnWidths: (sheet.get("widths") as Y.Map<number>).toJSON(),
    };
  }

  cellCount(name: string): number {
    return usedAddresses(this.cells(this.find(name))).size;
  }

  // The range from A1 to the last used cell, or null for an empty sheet.
  extent(name: string): string | null {
    const last = { col: 0, row: 0 };
    for (const addr of usedAddresses(this.cells(this.find(name)))) {
      const at = parseCell(addr);
      last.col = Math.max(last.col, at.col);
      last.row = Math.max(last.row, at.row);
    }
    return last.col === 0 ? null : `A1:${formatCell(last)}`;
  }

  setFrozenRows(name: string, rows: number): void {
    this.writable();
    if (!Number.isInteger(rows) || rows < 0 || rows > MAX_ROW) throw new SheetError(`invalid frozen rows ${rows}`);
    this.find(name).set("frozenRows", rows);
  }

  setColumnWidth(name: string, column: string, width: number): void {
    this.writable();
    const letters = column.toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(letters) || colIndex(letters) > MAX_COL) throw new SheetError(`invalid column "${column}"`);
    if (typeof width !== "number" || !Number.isFinite(width) || width <= 0) throw new SheetError(`invalid column width ${width}`);
    (this.find(name).get("widths") as Y.Map<number>).set(letters, width);
  }

  readRange(name: string, range: string): Cell[][] {
    const cells = this.cells(this.find(name));
    const span = this.span(range);
    return Array.from({ length: span.r2 - span.r1 + 1 }, (_, dr) =>
      Array.from({ length: span.c2 - span.c1 + 1 }, (_, dc) =>
        toCell(cells, formatCell({ col: span.c1 + dc, row: span.r1 + dr })),
      ),
    );
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
          const sheet = this.sheets.get(path[0] as number)?.get("name") as string | undefined;
          if (sheet === undefined) return listener({ all: true });
          for (const key of keys) if (key.endsWith(":c")) cells.push({ sheet, address: key.slice(0, -2) });
        } else if (path.length === 1 ? keys.includes("name") : path[1] !== "widths") {
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
    const content = this.cells(this.find(name)).get(`${formatCell({ col, row })}:c`);
    const valid = typeof content === "string" || typeof content === "boolean" || (typeof content === "number" && Number.isFinite(content));
    const formula = valid && isFormula(content) ? content : null;
    return { value: valid && formula === null ? (content as CellValue) : null, formula };
  }

  // Write a grid of values, or formulas, with its top-left corner at `anchor`.
  writeRange(name: string, anchor: string, rows: CellInput[][], by: Principal): void {
    this.writable();
    const cells = this.cells(this.find(name));
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
    this.doc.transact(() => {
      rows.forEach((line, dr) =>
        line.forEach((value, dc) => putValue(cells, formatCell({ col: start.col + dc, row: start.row + dr }), value, by)),
      );
    });
  }

  setFormat(name: string, range: string, patch: CellFormatPatch, by: Principal): void {
    this.writable();
    const cells = this.cells(this.find(name));
    const span = this.span(range, MAX_FORMAT_CELLS);
    checkPatch(patch);
    checkPrincipal(by);
    this.doc.transact(() => {
      for (let row = span.r1; row <= span.r2; row += 1) {
        for (let col = span.c1; col <= span.c2; col += 1) {
          const addr = formatCell({ col, row });
          for (const [field, value] of Object.entries(patch)) {
            if (value === null) cells.delete(`${addr}:fmt.${field}`);
            else if (value !== undefined) cells.set(`${addr}:fmt.${field}`, value);
          }
          tidyFormat(cells, addr, by);
        }
      }
    });
  }

  // Structure. `at` is the 1-based row or column. Refs to moved cells follow them. Refs to
  // deleted cells become #REF!. An insert that would push a cell or a width off the grid is refused.

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

  private restructure(name: string, axis: Axis, op: "insert" | "delete", at: number, count: number): void {
    this.writable();
    const sheet = this.find(name);
    const max = axis === "row" ? MAX_ROW : MAX_COL;
    if (!Number.isInteger(at) || !Number.isInteger(count) || at < 1 || at > max || count < 1 || count > MAX_STRUCTURE_COUNT) {
      throw new SheetError(`invalid ${axis} position ${at} or count ${count}`);
    }
    const cells = this.cells(sheet);
    const widths = sheet.get("widths") as Y.Map<number>;
    if (op === "insert") {
      let last = 0;
      for (const addr of usedAddresses(cells)) last = Math.max(last, parseCell(addr)[axis]);
      if (axis === "col") for (const letters of widths.keys()) last = Math.max(last, colIndex(letters));
      if (last >= at && last + count > max) throw new SheetError(`insert would push cells past the edge of the grid`);
    }
    const shift = (lo: number, hi: number): [number, number] | null => {
      const moved = shiftSpan(lo, hi, op, at, count);
      return moved === null || moved[0] > max ? null : [moved[0], Math.min(moved[1], max)];
    };
    this.doc.transact(() => {
      this.rewriteAll((ref) => {
        if (!sameSheet(ref.target, name)) return { range: ref.range };
        const { c1, r1, c2, r2 } = ref.range;
        const moved = axis === "row" ? shift(r1, r2) : shift(c1, c2);
        if (moved === null) return null;
        return { range: axis === "row" ? { c1, c2, r1: moved[0], r2: moved[1] } : { r1, r2, c1: moved[0], c2: moved[1] } };
      });
      const moves: { from: string; to: string | null; value: unknown }[] = [];
      for (const [from, value] of cells.entries()) {
        const cut = from.indexOf(":");
        const pos = parseCell(from.slice(0, cut));
        const moved = shift(pos[axis], pos[axis]);
        if (moved !== null && moved[0] === pos[axis]) continue;
        const to = moved === null ? null : `${formatCell({ ...pos, [axis]: moved[0] })}${from.slice(cut)}`;
        moves.push({ from, to, value });
      }
      for (const move of moves) cells.delete(move.from);
      for (const move of moves) if (move.to !== null) cells.set(move.to, move.value);
      if (axis === "col") this.moveWidths(widths, shift);
      else {
        const frozen = sheet.get("frozenRows") as number;
        if (op === "insert") sheet.set("frozenRows", at <= frozen ? Math.min(frozen + count, MAX_ROW) : frozen);
        else sheet.set("frozenRows", frozen - Math.max(0, Math.min(at + count - 1, frozen) - at + 1));
      }
      this.record();
    });
  }

  private moveWidths(widths: Y.Map<number>, shift: (lo: number, hi: number) => [number, number] | null): void {
    const entries = [...widths.entries()];
    for (const [letters] of entries) widths.delete(letters);
    for (const [letters, width] of entries) {
      const moved = shift(colIndex(letters), colIndex(letters));
      if (moved !== null) widths.set(colLetters(moved[0]), width);
    }
  }

  private rewriteAll(edit: RefEdit): void {
    for (const sheet of this.sheets) {
      const home = sheet.get("name") as string;
      const cells = this.cells(sheet);
      for (const [key, formula] of [...cells.entries()]) {
        if (!key.endsWith(":c") || !isFormula(formula)) continue;
        const next = rewriteRefs(formula, home, edit);
        if (next !== formula) cells.set(key, next);
      }
    }
  }

  private find(name: string): YSheet {
    const sheet = this.sheets.toArray().find((s) => sameSheet(s.get("name") as string, name));
    if (sheet === undefined) throw new SheetError(`no sheet "${name}"`);
    return sheet;
  }

  private cells(sheet: YSheet): Cells {
    return sheet.get("cells") as Cells;
  }

  private checkName(name: string): void {
    if (name.trim() === "" || BAD_SHEET_NAME.test(name)) throw new SheetError(`invalid sheet name "${name}"`);
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

function checkPrincipal(by: Principal): void {
  if (by.id === "" || !PRINCIPAL_KINDS.includes(by.kind)) throw new SheetError("invalid principal");
}

const isFormula = (v: unknown): v is string => typeof v === "string" && v.startsWith("=") && v.length > 1;

function checkValue(value: unknown): void {
  if (value === null || typeof value === "boolean") return;
  if (typeof value === "number" && Number.isFinite(value)) return;
  if (typeof value !== "string") throw new SheetError("invalid cell value");
  if (isFormula(value) && value.length > MAX_FORMULA) throw new SheetError(`formula longer than ${MAX_FORMULA} characters`);
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

// The distinct cell addresses that hold any key.
function usedAddresses(cells: Cells): Set<string> {
  const used = new Set<string>();
  for (const key of cells.keys()) used.add(key.slice(0, key.indexOf(":")));
  return used;
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

function toCell(cells: Cells, addr: string): Cell {
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
  const formula = valid && isFormula(content) ? content : null;
  return {
    value: valid && formula === null ? (content as CellValue) : null,
    formula,
    format: format as CellFormat,
    author: toAuthor(by),
    formatAuthor: toAuthor(fby),
  };
}
