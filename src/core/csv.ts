import type { CellInput } from "./workbook.js";

// Split RFC 4180 text into rows. Quoted fields may hold commas, quotes, and newlines.
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  let seen = false;
  for (let i = text.startsWith("\uFEFF") ? 1 : 0; i < text.length; i += 1) {
    const ch = text[i];
    if (quoted) {
      if (ch !== '"') field += ch;
      else if (text[i + 1] === '"') (field += '"', (i += 1));
      else quoted = false;
    } else if (ch === '"') (quoted = true), (seen = true);
    else if (ch === ",") (row.push(field), (field = ""));
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i += 1;
      row.push(field);
      rows.push(row);
      [row, field, seen] = [[], "", false];
    } else field += ch;
  }
  if (field !== "" || row.length > 0 || seen) rows.push([...row, field]);
  return rows;
}

// Text that starts with = + - @ could run as a formula in a spreadsheet
// app, so export marks it with a leading quote. Numbers are not marked.
export function formatCsv(rows: unknown[][]): string {
  const cell = (v: unknown): string => {
    let text = v === null || v === undefined ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v);
    if (typeof v === "object" && v !== null && "error" in v) text = String(v.error);
    if (typeof v === "string" && /^[=+\-@\t\r]/.test(v)) text = `'${v}`;
    return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };
  return rows.map((row) => `${row.map(cell).join(",")}\r\n`).join("");
}

// Numbers must be plain, so "007" stays text, and so does an id with more than 15 digits.
// A leading ' on text that starts with + - @ or a control character is the mark that export
// adds, so it is removed. A leading = makes a formula, unless `asText` keeps it as text.
export function inferCell(text: string, asText = false): CellInput {
  if (text === "") return null;
  if (/^'[+\-@\t\r]/.test(text)) return text.slice(1);
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(text) && text.replace(/\D/g, "").length <= 15) return Number(text);
  if (asText && text.startsWith("=")) return `'${text}`;
  return text === "TRUE" ? true : text === "FALSE" ? false : text;
}
