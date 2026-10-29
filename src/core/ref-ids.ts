import type { Flags } from "./refs.js";

// A stored formula keeps each cell ref as a token that names a sheet, a row, and a column by id.
// The token starts with U+0001 and ends with U+0002. A formula typed by a person may not hold
// either character, so a token never comes from outside.
//
//   U+0001 sheet | row1 | col1 | row2 | col2 | bits U+0002
//
// The bits are q (the ref names its sheet), R (a range), and a b c d (a $ before col1, row1, col2, row2).

export interface IdRef {
  sheet: string;
  r1: string;
  c1: string;
  r2: string;
  c2: string;
  qualified: boolean;
  flags: Flags;
}

const TOKEN = /\u0001([^\u0002]*)\u0002/g;

export const hasControl = (text: string): boolean => /[\u0001\u0002]/.test(text);

export function packRef({ sheet, r1, c1, r2, c2, qualified, flags }: IdRef): string {
  const bits = `${qualified ? "q" : ""}${flags.isRange ? "R" : ""}${flags.c1 ? "a" : ""}${flags.r1 ? "b" : ""}${flags.c2 ? "c" : ""}${flags.r2 ? "d" : ""}`;
  return `\u0001${sheet}|${r1}|${c1}|${r2}|${c2}|${bits}\u0002`;
}

// Replace each token with the text `show` returns. A token that does not parse reads as #REF!.
export function unpackRefs(stored: string, show: (ref: IdRef) => string): string {
  return stored.replace(TOKEN, (_, body: string) => {
    const [sheet, r1, c1, r2, c2, bits] = body.split("|");
    if (bits === undefined || [sheet, r1, c1, r2, c2].some((part) => part === "")) return "#REF!";
    const has = (bit: string) => bits.includes(bit);
    const flags: Flags = { c1: has("a") ? "$" : "", r1: has("b") ? "$" : "", c2: has("c") ? "$" : "", r2: has("d") ? "$" : "", isRange: has("R") };
    return show({ sheet, r1, c1, r2, c2, qualified: has("q"), flags });
  });
}

// Apply `edit` to the parts of a stored formula that are not tokens. A token holds ids that could
// read as a cell address, so no edit may see one.
export function mapPlain(stored: string, edit: (plain: string) => string): string {
  let out = "";
  let at = 0;
  for (const match of stored.matchAll(TOKEN)) {
    out += edit(stored.slice(at, match.index)) + match[0];
    at = match.index + match[0].length;
  }
  return out + edit(stored.slice(at));
}
