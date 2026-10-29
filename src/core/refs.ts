import { colIndex, colLetters, MAX_COL, MAX_ROW, type RangePos } from "./a1.js";

// A cell or range reference found in a formula.
export interface FoundRef {
  // The sheet name as written, or null for a ref on the formula's own sheet.
  sheet: string | null;
  // The sheet the ref points at.
  target: string;
  range: RangePos;
}

const REF =
  /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?(?![\w(])/y;

export function quoteSheet(name: string): string {
  const plain = /^[A-Za-z_][\w.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name);
  return plain ? name : `'${name.replaceAll("'", "''")}'`;
}

export function sameSheet(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

function skipLiteral(text: string, start: number): number {
  let i = start + 1;
  while (i < text.length) {
    if (text[i] !== '"') i += 1;
    else if (text[i + 1] === '"') i += 2;
    else break;
  }
  return i + 1;
}

// Replace every ref in a formula with the text `edit` returns. Text inside string literals, and
// words such as LOG10( that only look like refs, stay as they are.
export function mapRefs(formula: string, home: string, edit: (found: { text: string; ref: FoundRef; flags: Flags }) => string): string {
  let out = "";
  let i = 0;
  while (i < formula.length) {
    if (formula[i] === '"') {
      const end = skipLiteral(formula, i);
      out += formula.slice(i, end);
      i = end;
      continue;
    }
    const match = /[\w.$!]/.test(formula[i - 1] ?? "") ? null : readRef(formula, i, home);
    if (match === null) {
      out += formula[i];
      i += 1;
      continue;
    }
    out += edit(match);
    i += match.text.length;
  }
  return out;
}

export interface Flags {
  c1: string;
  r1: string;
  c2: string;
  r2: string;
  isRange: boolean;
}

export function readRef(formula: string, at: number, home: string): { text: string; ref: FoundRef; flags: Flags } | null {
  REF.lastIndex = at;
  const m = REF.exec(formula);
  if (m === null) return null;
  const [text, quoted, bare, ac1, l1, ar1, n1, ac2, l2, ar2, n2] = m;
  const c1 = colIndex(l1);
  const r1 = Number(n1);
  const c2 = l2 === undefined ? c1 : colIndex(l2);
  const r2 = n2 === undefined ? r1 : Number(n2);
  if (Math.max(c1, c2) > MAX_COL || Math.min(r1, r2) < 1 || Math.max(r1, r2) > MAX_ROW) return null;
  const sheet = quoted !== undefined ? quoted.replaceAll("''", "'") : (bare ?? null);
  const range = {
    c1: Math.min(c1, c2),
    r1: Math.min(r1, r2),
    c2: Math.max(c1, c2),
    r2: Math.max(r1, r2),
  };
  const flags = { c1: ac1, r1: ar1, c2: ac2 ?? "", r2: ar2 ?? "", isRange: l2 !== undefined };
  return { text, ref: { sheet, target: sheet ?? home, range }, flags };
}

export function formatSpan(range: RangePos, flags: Flags): string {
  const start = `${flags.c1}${colLetters(range.c1)}${flags.r1}${range.r1}`;
  if (!flags.isRange) return start;
  return `${start}:${flags.c2}${colLetters(range.c2)}${flags.r2}${range.r2}`;
}
