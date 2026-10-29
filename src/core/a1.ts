export const MAX_COL = 16384;
export const MAX_ROW = 1048576;

export interface CellPos {
  col: number;
  row: number;
}

export interface RangePos {
  c1: number;
  r1: number;
  c2: number;
  r2: number;
}

export function colIndex(letters: string): number {
  let index = 0;
  for (const char of letters.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index;
}

export function colLetters(index: number): string {
  let letters = "";
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

export function parseCell(text: string): CellPos {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(text);
  const col = match ? colIndex(match[1]) : 0;
  const row = match ? Number(match[2]) : 0;
  if (!match || col > MAX_COL || row < 1 || row > MAX_ROW) {
    throw new Error(`invalid cell "${text}"`);
  }
  return { col, row };
}

export function formatCell(pos: CellPos): string {
  return `${colLetters(pos.col)}${pos.row}`;
}

export function parseRange(text: string): RangePos {
  const parts = text.split(":");
  if (parts.length > 2 || parts.some((part) => part === "")) {
    throw new Error(`invalid range "${text}"`);
  }
  const a = parseCell(parts[0]);
  const b = parseCell(parts[1] ?? parts[0]);
  return {
    c1: Math.min(a.col, b.col),
    r1: Math.min(a.row, b.row),
    c2: Math.max(a.col, b.col),
    r2: Math.max(a.row, b.row),
  };
}
