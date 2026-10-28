import * as Y from "yjs";
import { colIndex, formatCell, MAX_COL, MAX_ROW, parseCell, parseRange, type RangePos } from "./a1.js";

export type PrincipalKind = "human" | "agent_of_human" | "agent_of_agent" | "agent" | "user_token";

const PRINCIPAL_KINDS: readonly PrincipalKind[] = ["human", "agent_of_human", "agent_of_agent", "agent", "user_token"];

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
  author: Author | null;
}

export interface SheetInfo {
  name: string;
  frozenRows: number;
  columnWidths: Record<string, number>;
}

export class SheetError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SheetError";
  }
}

const MAX_RANGE_CELLS = 1_000_000;
const BAD_SHEET_NAME = /[[\]:*?/\\!']/;

type YCell = Y.Map<unknown>;
type YSheet = Y.Map<unknown>;

function sameSheet(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

export class Workbook {
  readonly doc: Y.Doc;
  private readonly sheets: Y.Array<YSheet>;

  constructor(doc: Y.Doc = new Y.Doc()) {
    this.doc = doc;
    this.sheets = doc.getArray("sheets");
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
    this.checkName(name);
    const sheet: YSheet = new Y.Map();
    sheet.set("name", name);
    sheet.set("cells", new Y.Map());
    sheet.set("widths", new Y.Map());
    sheet.set("frozenRows", 0);
    this.sheets.insert(Math.max(0, Math.min(index, this.sheets.length)), [sheet]);
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
    return this.cells(this.find(name)).size;
  }

  setFrozenRows(name: string, rows: number): void {
    if (!Number.isInteger(rows) || rows < 0 || rows > MAX_ROW) throw new SheetError(`invalid frozen rows ${rows}`);
    this.find(name).set("frozenRows", rows);
  }

  setColumnWidth(name: string, column: string, width: number): void {
    const letters = column.toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(letters) || colIndex(letters) > MAX_COL) throw new SheetError(`invalid column "${column}"`);
    if (!(width > 0)) throw new SheetError(`invalid column width ${width}`);
    (this.find(name).get("widths") as Y.Map<number>).set(letters, width);
  }

  readRange(name: string, range: string): Cell[][] {
    const cells = this.cells(this.find(name));
    const span = this.span(range);
    return Array.from({ length: span.r2 - span.r1 + 1 }, (_, dr) =>
      Array.from({ length: span.c2 - span.c1 + 1 }, (_, dc) =>
        toCell(cells.get(formatCell({ col: span.c1 + dc, row: span.r1 + dr }))),
      ),
    );
  }

  // Write a grid of values, or formulas, with its top-left corner at `anchor`.
  writeRange(name: string, anchor: string, rows: CellInput[][], by: Principal): void {
    const cells = this.cells(this.find(name));
    const start = parseCell(anchor);
    const width = Math.max(0, ...rows.map((line) => line.length));
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
    const cells = this.cells(this.find(name));
    const span = this.span(range);
    checkPrincipal(by);
    this.doc.transact(() => {
      for (let row = span.r1; row <= span.r2; row += 1) {
        for (let col = span.c1; col <= span.c2; col += 1) {
          const key = formatCell({ col, row });
          const cell = (cells.get(key) as YCell | undefined) ?? (cells.set(key, new Y.Map()) as YCell);
          const format: Record<string, unknown> = { ...(cell.get("fmt") as object | undefined), ...patch };
          for (const [field, value] of Object.entries(format)) if (value === null) delete format[field];
          cell.set("fmt", format);
          stamp(cell, by);
        }
      }
    });
  }

  private find(name: string): YSheet {
    const sheet = this.sheets.toArray().find((s) => sameSheet(s.get("name") as string, name));
    if (sheet === undefined) throw new SheetError(`no sheet "${name}"`);
    return sheet;
  }

  private cells(sheet: YSheet): Y.Map<YCell> {
    return sheet.get("cells") as Y.Map<YCell>;
  }

  private checkName(name: string): void {
    if (name.trim() === "" || BAD_SHEET_NAME.test(name)) throw new SheetError(`invalid sheet name "${name}"`);
    if (this.sheetNames().some((existing) => sameSheet(existing, name))) {
      throw new SheetError(`sheet "${name}" exists`);
    }
  }

  private span(range: string): RangePos {
    const span = parseRange(range);
    if ((span.c2 - span.c1 + 1) * (span.r2 - span.r1 + 1) > MAX_RANGE_CELLS) {
      throw new SheetError(`range ${range} is too large`);
    }
    return span;
  }
}

function checkPrincipal(by: Principal): void {
  if (by.id === "" || !PRINCIPAL_KINDS.includes(by.kind)) throw new SheetError("invalid principal");
}

function stamp(cell: YCell, by: Principal): void {
  cell.set("by", { id: by.id, kind: by.kind });
  cell.set("at", new Date().toISOString());
}

function putValue(cells: Y.Map<YCell>, key: string, value: CellInput, by: Principal): void {
  const existing = cells.get(key);
  if (value === null || value === "") {
    if (existing === undefined) return;
    if (!existing.has("fmt")) {
      cells.delete(key);
      return;
    }
    existing.delete("v");
    existing.delete("f");
    stamp(existing, by);
    return;
  }
  const cell = existing ?? (cells.set(key, new Y.Map()) as YCell);
  if (typeof value === "string" && value.startsWith("=") && value.length > 1) {
    cell.set("f", value);
    cell.delete("v");
  } else {
    cell.set("v", value);
    cell.delete("f");
  }
  stamp(cell, by);
}

function toCell(cell: YCell | undefined): Cell {
  if (cell === undefined) return { value: null, formula: null, format: {}, author: null };
  const by = cell.get("by") as Principal | undefined;
  return {
    value: (cell.get("v") as CellValue | undefined) ?? null,
    formula: (cell.get("f") as string | undefined) ?? null,
    format: { ...(cell.get("fmt") as CellFormat | undefined) },
    author: by === undefined ? null : { ...by, machine: isMachineMade(by.kind), at: cell.get("at") as string },
  };
}
