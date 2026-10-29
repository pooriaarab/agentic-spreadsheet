import * as Y from "yjs";
import { DEFAULT_LIMITS, NUMBER_FORMATS, ODS_MIME, parseRef, readOds, readXlsx, refOf, writeOds, writeXlsx } from "agentic-formats";
import type { Cell as FileCell, Workbook as FileBook } from "agentic-formats";
import { parseRange } from "./core/a1.js";
import { inferCell } from "./core/csv.js";
import { Engine } from "./core/engine.js";
import { PRINCIPAL_KINDS, SheetError, Workbook, type CellInput, type Principal } from "./core/workbook.js";

export type SheetFormat = "ods" | "xlsx";
export const MIME: Record<SheetFormat, string> = {
  ods: ODS_MIME,
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

/** The format a file name asks for, by its extension, or null. */
export function formatOf(fileName: string): SheetFormat | null {
  const parts = fileName.toLowerCase().split(".");
  const ext = parts.length > 1 ? parts.pop() : undefined;
  return ext === "ods" || ext === "xlsx" ? ext : null;
}

const numberFormats = new Set<string>(NUMBER_FORMATS);

/** The workbook as a neutral one. A formula cell carries the value the engine gives it. */
export function toFileBook(book: Workbook): FileBook {
  const engine = new Engine(book);
  // The cost follows the grids, so the limit counts the used ranges of all the sheets.
  let grid = 0;
  try {
    const sheets = book.sheetNames().map((name) => {
      const cells: Record<string, FileCell> = {};
      const extent = book.extent(name);
      if (extent) {
        const { c2, r2 } = parseRange(extent);
        grid += c2 * r2;
        if (grid > DEFAULT_LIMITS.maxCells) throw new SheetError(`the sheet ${name} is too large to export`);
        const read = book.readRange(name, extent);
        const values = engine.values(name, extent);
        read.forEach((row, r) =>
          row.forEach((cell, c) => {
            const out: FileCell = {};
            if (cell.formula) {
              out.formula = cell.formula.replace(/^=/, "");
              const shown = values[r]?.[c];
              if (typeof shown === "string" || typeof shown === "number" || typeof shown === "boolean") out.value = shown;
            } else if (cell.value !== null) out.value = typeof cell.value === "string" && cell.value.startsWith("'=") ? cell.value.slice(1) : cell.value;
            const numberFormat = cell.format.numberFormat;
            if (numberFormat && numberFormats.has(numberFormat)) out.format = numberFormat;
            if (out.value !== undefined || out.formula) cells[refOf(c + 1, r + 1)] = out;
          }),
        );
      }
      return { name, cells };
    });
    return { sheets };
  } finally {
    engine.dispose();
  }
}

/** The workbook as the bytes of an ODS or XLSX file, with the notes of what the format could not hold. */
export function exportFile(format: SheetFormat, book: Workbook): { bytes: Uint8Array; warnings: string[] } {
  const out = format === "ods" ? writeOds(toFileBook(book)) : writeXlsx(toFileBook(book));
  return { bytes: out.bytes, warnings: out.warnings };
}

/** The most sheets one import adds. */
export const MAX_IMPORT_SHEETS = 100;

/** The longest sheet name that XLSX holds. */
const MAX_NAME = 31;

/** Names are compared without case, as the workbook does. A name is cut to the XLSX maximum. */
function freeName(taken: Set<string>, requested: string): string {
  const wanted = requested.slice(0, MAX_NAME);
  if (!taken.has(wanted.toLowerCase())) return wanted;
  for (let n = 2; ; n++) {
    const next = `${wanted.slice(0, MAX_NAME - ` (${n})`.length)} (${n})`;
    if (!taken.has(next.toLowerCase())) return next;
  }
}

/**
 * The name each sheet gets in the workbook. A name that is taken, or has a banned character, changes. A formula names
 * a sheet as text, and the text would then point at another sheet, or at none, so a file with a rename and a formula that names a sheet is refused.
 */
function renamed(model: FileBook, taken: Set<string>): string[] {
  const names = model.sheets.map((sheet) => {
    const name = freeName(taken, sheet.name.replace(/[[\]:*?/\\!']/g, "_"));
    taken.add(name.toLowerCase());
    return name;
  });
  // A sheet name sits before `!`, in plain or quoted form (`'Bob''s'!A1`), so any formula with `!` may name a changed sheet.
  const crossSheet = model.sheets.some((sheet) => Object.values(sheet.cells).some((cell) => cell.formula?.includes("!")));
  const changed = model.sheets.findIndex((sheet, i) => names[i] !== sheet.name);
  if (crossSheet && changed >= 0) {
    throw new SheetError(`the sheet "${model.sheets[changed]!.name}" would be renamed "${names[changed]}", and a formula in the file names a sheet`);
  }
  return names;
}

function writeSheets(book: Workbook, model: FileBook, by: Principal): string[] {
  if (model.sheets.length > MAX_IMPORT_SHEETS) throw new SheetError(`a file can add up to ${MAX_IMPORT_SHEETS} sheets at a time`);
  if (by.id === "" || !PRINCIPAL_KINDS.includes(by.kind)) throw new SheetError("invalid principal");
  let grid = 0;
  const taken = new Set(book.sheetNames().map((n) => n.toLowerCase()));
  const names = renamed(model, taken);
  for (const [i, sheet] of model.sheets.entries()) {
    const name = names[i]!;
    book.addSheet(name);
    let rows = 0;
    let cols = 0;
    const placed: [number, number, FileCell][] = [];
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      const at = parseRef(ref);
      if (!at) continue;
      rows = Math.max(rows, at.row);
      cols = Math.max(cols, at.column);
      placed.push([at.column, at.row, cell]);
    }
    if (!placed.length) continue;
    // The grid is as big as its far corner, so a file with two far-apart cells cannot make a huge one.
    // The cost follows the grids, not the cells, so the limit counts the grids of all the sheets.
    grid += rows * cols;
    if (grid > DEFAULT_LIMITS.maxCells) throw new SheetError(`the sheet ${sheet.name} is too large to import`);
    const matrix: CellInput[][] = Array.from({ length: rows }, () => Array<CellInput>(cols).fill(null));
    for (const [c, r, cell] of placed) {
      matrix[r - 1]![c - 1] = cell.formula ? `=${cell.formula}` : typeof cell.value === "string" && cell.value.startsWith("=") && cell.value.length > 1 ? inferCell(cell.value, true) : (cell.value ?? null);
    }
    book.writeRange(name, "A1", matrix, by);
    for (const [c, r, cell] of placed) if (cell.format) book.setFormat(name, refOf(c, r), { numberFormat: cell.format }, by);
  }
  return names;
}

/**
 * Runs `work` on a scratch copy of the document. Only when it returns does the change reach the real document, so
 * an error leaves the workbook as it was. A Yjs transaction alone does not roll back.
 */
function atomically<T>(doc: Y.Doc, work: (scratch: Y.Doc) => T): T {
  const scratch = new Y.Doc();
  Y.applyUpdate(scratch, Y.encodeStateAsUpdate(doc));
  try {
    const result = work(scratch);
    Y.applyUpdate(doc, Y.encodeStateAsUpdate(scratch, Y.encodeStateVector(doc)));
    return result;
  } finally {
    scratch.destroy();
  }
}

/** Adds the sheets of a neutral workbook after the sheets that are there. A taken name gets a number. All of the sheets are added, or none. */
export const addSheets = (book: Workbook, model: FileBook, by: Principal): string[] => atomically(book.doc, (scratch) => writeSheets(new Workbook(scratch), model, by));

/**
 * Reads the bytes of an ODS or XLSX file and adds its sheets to the workbook. The input is untrusted: a file over
 * the size limit, a file with no sheets, or a grid over the cell limit throws a SheetError and leaves the workbook as it was.
 * A file the reader refuses throws the FormatError of agentic-formats.
 */
export function importFile(format: SheetFormat, bytes: Uint8Array, book: Workbook, by: Principal): { names: string[]; warnings: string[] } {
  if (bytes.byteLength > DEFAULT_LIMITS.maxInputBytes) throw new SheetError(`files can be up to ${DEFAULT_LIMITS.maxInputBytes / 1024 / 1024} MB`);
  const read = format === "ods" ? readOds(bytes) : readXlsx(bytes);
  if (!read.model.sheets.length) throw new SheetError("the file has no sheets to import");
  return { names: addSheets(book, read.model, by), warnings: read.warnings };
}
