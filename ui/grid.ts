import { colIndex, colLetters } from "../src/core/a1.js";
import type { Cell, FormulaError, Value } from "../src/client.js";

export interface Pos {
  row: number;
  col: number;
}
export const MIN_COLS = 10;
export const MIN_ROWS = 30;
const MAX_COLS = 52;
const MAX_ROWS = 200;

// Sizes in pixels.
export const COL_WIDTH = 112;
export const ROW_HEIGHT = 28;
export const ROW_HEAD = 48;
const LIMITS = { cols: [40, 640], rows: [22, 240] } as const;
export const clampSize = (axis: "cols" | "rows", size: number): number => Math.round(Math.min(LIMITS[axis][1], Math.max(LIMITS[axis][0], size)));

export const address = ({ row, col }: Pos): string => `${colLetters(col)}${row}`;

// The size of the grid to show: the used range plus room to grow.
export function sizeFor(extent: string | null): { cols: number; rows: number } {
  const end = extent?.split(":").pop()?.match(/^([A-Z]+)(\d+)$/);
  if (!end) return { cols: MIN_COLS, rows: MIN_ROWS };
  return { cols: Math.min(MAX_COLS, Math.max(MIN_COLS, colIndex(end[1]) + 3)), rows: Math.min(MAX_ROWS, Math.max(MIN_ROWS, Number(end[2]) + 10)) };
}

// Text from the editor to what the API stores: "=..." is a formula, a plain number is a number, empty clears.
export function parseInput(text: string): Cell {
  if (text === "") return null;
  if (!text.startsWith("=") && text.trim() !== "" && Number.isFinite(Number(text))) return Number(text);
  return text;
}

export const isError = (value: Value): value is FormulaError => typeof value === "object" && value !== null;

export function show(value: Value): string {
  if (value === null) return "";
  if (isError(value)) return value.error;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

// A peer sets its own cursor. Only an A1 address passes.
export const cursorOf = (value: unknown): string | null => (typeof value === "string" && /^[A-Z]{1,3}[1-9]\d{0,6}$/.test(value) ? value : null);
