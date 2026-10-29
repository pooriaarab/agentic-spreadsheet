import { z } from "zod";
import { colLetters, MAX_ROW, parseRange } from "./core/a1.js";
import { Engine } from "./core/engine.js";
import { formatCsv, inferCell, parseCsv } from "./core/csv.js";
import { fromBase64, toBase64 } from "./core/bytes.js";
import type { Store } from "./core/store.js";
import { MAX_STRUCTURE_COUNT, Workbook, type Principal } from "./core/workbook.js";

// One list of operations serves the CLI and the MCP server.

export interface Ctx {
  store: Store;
  // Who is writing. It throws when the caller did not say, so no edit gets a default label.
  principal(): Principal;
  // The most cells one write or import may carry. The CLI sets none. The Worker sets a cap.
  maxWriteCells?: number;
}

// An error with an HTTP-style status, so the Worker can answer 404 or 409 and not a blanket 400.
export class OpError extends Error {
  constructor(
    message: string,
    readonly status: 404 | 409 | 413,
  ) {
    super(message);
  }
}

export function message(error: unknown): string {
  if (error instanceof z.ZodError) return error.issues.map((i) => `${i.path.join(".")}: ${i.message}`.replace(/^: /, "")).join("; ");
  return error instanceof Error ? error.message : String(error);
}

export interface Op {
  name: string;
  description: string;
  // Positional argument names, for the CLI.
  args: string[];
  shape: z.ZodRawShape;
  readOnly?: boolean;
  run(args: any, ctx: Ctx): Promise<unknown>;
}

const book = z.string().regex(/^[\w-][\w.-]*$/, "must use letters, digits, dot, dash, or underscore").describe("Workbook id.");
const sheet = z.string().optional().describe("Sheet name. Defaults to the first sheet.");
const grid = z.preprocess(
  (v) => {
    try {
      return typeof v === "string" ? JSON.parse(v) : v;
    } catch {
      return v;
    }
  },
  z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))),
).describe('Rows of values. A string that starts with "=" is a formula.');

async function open(store: Store, id: string): Promise<Workbook> {
  const doc = await store.get("workbooks", id);
  if (doc === null) throw new OpError(`no workbook "${id}"`, 404);
  return Workbook.fromState(fromBase64(doc.state as string));
}

// Read, change, and save one workbook while holding the store lock. The workbook is saved only
// when `change` returns, so a failed edit leaves the stored copy as it was.
function edit<T>(ctx: Ctx, id: string, change: (wb: Workbook) => T): Promise<T> {
  return ctx.store.transaction(async (store) => {
    const wb = await open(store, id);
    const result = change(wb);
    await store.put("workbooks", id, { name: id, state: toBase64(wb.encode()), updatedAt: new Date().toISOString() });
    return result;
  });
}

function checkWriteSize(ctx: Ctx, rows: unknown[][]): void {
  const cells = rows.reduce<number>((sum, row) => sum + row.length, 0);
  if (ctx.maxWriteCells !== undefined && cells > ctx.maxWriteCells) {
    throw new OpError(`a write of ${cells.toLocaleString("en-US")} cells is too large (limit ${ctx.maxWriteCells.toLocaleString("en-US")} cells)`, 413);
  }
}

const first = (wb: Workbook, name?: string) => name ?? wb.sheetNames()[0];
// Export reads the sheet in blocks of this many cells. It refuses a sheet whose extent is larger
// than the total cap, because the CSV holds every cell from A1 to the last used one.
const EXPORT_BLOCK_CELLS = 100_000;
const MAX_EXPORT_CELLS = 1_000_000;
export const OPS: Op[] = [
  {
    name: "new", description: "Create a workbook with one empty sheet.", args: ["book"],
    shape: { book, sheet: z.string().optional().describe("Name of the first sheet. Defaults to Sheet1.") },
    run(a: { book: string; sheet?: string }, ctx) {
      return ctx.store.transaction(async (store) => {
        if ((await store.get("workbooks", a.book)) !== null) throw new OpError(`workbook "${a.book}" exists`, 409);
        const wb = new Workbook();
        wb.addSheet(a.sheet ?? "Sheet1");
        await store.put("workbooks", a.book, { name: a.book, state: toBase64(wb.encode()), updatedAt: new Date().toISOString() });
        return { book: a.book, sheets: wb.sheetNames() };
      });
    },
  },
  {
    name: "list", description: "List workbook ids in name order.", args: [], readOnly: true,
    shape: { limit: z.coerce.number().int().min(1).optional(), cursor: z.string().optional() },
    async run(a: { limit?: number; cursor?: string }, ctx) {
      const { items, cursor } = await ctx.store.query("workbooks", a);
      return { books: items.map((i) => i.id), cursor };
    },
  },
  {
    name: "read", description: "Read a range: computed values, formulas, and who last edited each cell.", args: ["book", "range"], readOnly: true,
    shape: { book, range: z.string().describe("A1 range, such as A1:C5."), sheet },
    async run(a: { book: string; range: string; sheet?: string }, ctx) {
      const wb = await open(ctx.store, a.book);
      const name = first(wb, a.sheet);
      const cells = wb.readRange(name, a.range);
      return {
        sheet: name, range: a.range, values: new Engine(wb).values(name, a.range),
        formulas: cells.map((row) => row.map((c) => c.formula)), authors: cells.map((row) => row.map((c) => c.author)),
      };
    },
  },
  {
    name: "write", description: "Write a grid of values or formulas from a start cell.", args: ["book", "anchor", "values"],
    shape: { book, anchor: z.string().describe("Top-left cell, such as A1."), values: grid, sheet },
    async run(a: { book: string; anchor: string; values: (string | number | boolean | null)[][]; sheet?: string }, ctx) {
      const who = ctx.principal();
      checkWriteSize(ctx, a.values);
      await edit(ctx, a.book, (wb) => wb.writeRange(first(wb, a.sheet), a.anchor, a.values, who));
      return { ok: true, rows: a.values.length };
    },
  },
  {
    name: "formula", description: "Set one cell to a formula and return its value.", args: ["book", "cell", "formula"],
    shape: { book, cell: z.string().describe("Cell, such as B2."), formula: z.string().describe("Formula text. The leading = is optional."), sheet },
    async run(a: { book: string; cell: string; formula: string; sheet?: string }, ctx) {
      const who = ctx.principal();
      const formula = a.formula.startsWith("=") ? a.formula : `=${a.formula}`;
      return edit(ctx, a.book, (wb) => {
        const name = first(wb, a.sheet);
        wb.writeRange(name, a.cell, [[formula]], who);
        return { cell: a.cell, formula, value: new Engine(wb).value(name, a.cell) };
      });
    },
  },
  {
    name: "insert-rows", description: "Insert rows. Formulas follow the moved cells.", args: ["book", "at", "count"],
    shape: {
      book,
      at: z.coerce.number().int().min(1).max(MAX_ROW).describe("Row number to insert before."),
      count: z.coerce.number().int().min(1).max(MAX_STRUCTURE_COUNT).default(1),
      sheet,
    },
    async run(a: { book: string; at: number; count: number; sheet?: string }, ctx) {
      ctx.principal();
      await edit(ctx, a.book, (wb) => wb.insertRows(first(wb, a.sheet), a.at, a.count));
      return { ok: true };
    },
  },
  {
    name: "export", description: "Export a sheet as CSV with computed values.", args: ["book"], readOnly: true,
    shape: { book, sheet },
    async run(a: { book: string; sheet?: string }, ctx) {
      const wb = await open(ctx.store, a.book);
      const name = first(wb, a.sheet);
      const extent = wb.extent(name);
      if (extent === null) return { csv: "" };
      const { c2, r2 } = parseRange(extent);
      if (c2 * r2 > MAX_EXPORT_CELLS) {
        throw new OpError(`sheet ${name} spans ${extent}, which is too large to export (limit ${MAX_EXPORT_CELLS.toLocaleString("en-US")} cells)`, 413);
      }
      // Read the sheet in blocks sized by cells, so a wide sheet never builds one giant grid.
      const rowsPerBlock = Math.max(1, Math.floor(EXPORT_BLOCK_CELLS / c2));
      const engine = new Engine(wb);
      const parts: string[] = [];
      for (let row = 1; row <= r2; row += rowsPerBlock) {
        parts.push(formatCsv(engine.values(name, `A${row}:${colLetters(c2)}${Math.min(row + rowsPerBlock - 1, r2)}`)));
      }
      return { csv: parts.join("") };
    },
  },
  {
    name: "import", description: "Write CSV text into a sheet from A1.", args: ["book"],
    shape: {
      book,
      csv: z.string().describe("CSV text."),
      asText: z.boolean().default(false).describe("Keep a cell that starts with = as text. Use this for CSV from an untrusted source."),
      sheet,
    },
    async run(a: { book: string; csv: string; asText: boolean; sheet?: string }, ctx) {
      const who = ctx.principal();
      const rows = parseCsv(a.csv).map((row) => row.map((text) => inferCell(text, a.asText)));
      checkWriteSize(ctx, rows);
      await edit(ctx, a.book, (wb) => wb.writeRange(first(wb, a.sheet), "A1", rows, who));
      return { ok: true, rows: rows.length };
    },
  },
];
