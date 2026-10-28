import * as formulajs from "@formulajs/formulajs";
import { formatCell, parseCell, parseRange, type RangePos } from "./a1.js";
import { DepthError, parse, ParseError, type Node } from "./formula.js";
import { sameSheet } from "./refs.js";
import { MAX_TEXT, type Workbook } from "./workbook.js";

export type ErrorCode = "#REF!" | "#DIV/0!" | "#NAME?" | "#VALUE!" | "#NUM!" | "#N/A" | "#CIRCULAR" | "#DEPTH!";
export type ErrorValue = { error: ErrorCode };
export type Value = number | string | boolean | null | ErrorValue;
type Matrix = Value[][];
type Ctx = { home: string };
type Ref = Extract<Node, { t: "ref" }>;

const ERRORS = new Set<string>(["#REF!", "#DIV/0!", "#NAME?", "#VALUE!", "#NUM!", "#N/A", "#CIRCULAR", "#DEPTH!"]);
const MAX_RANGE_CELLS = 1_000_000;
// Range cells one cell may read, and levels of cell-to-cell nesting one read may follow.
const MAX_WORK = 10_000_000;
const MAX_DEPTH = 400;
// formulajs reads a blank cell as 0. These groups need another reading.
const TEXT_FUNCTIONS = new Set([
  "LEFT", "RIGHT", "MID", "LEN", "UPPER", "LOWER", "TRIM", "CONCAT", "TEXT", "CONCATENATE", "SUBSTITUTE", "FIND",
  "SEARCH", "EXACT", "REPLACE", "PROPER", "CLEAN", "REPT", "TEXTJOIN",
]);
const BLANK_CELLS = new Set([
  "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS", "MAXIFS", "MINIFS", "COUNTBLANK", "MATCH",
  "VLOOKUP", "HLOOKUP", "LOOKUP", "INDEX",
]);
// Statistics that skip blanks, text, and booleans in a range. Variadic ones also take refs one by one.
const VARIADIC = new Set([
  "PRODUCT", "SUMSQ", "MEDIAN", "STDEVS", "STDEVP", "VARS", "VARP", "GEOMEAN", "HARMEAN", "AVEDEV", "DEVSQ", "KURT", "SKEW", "SKEWP",
]);
const SKIP_BLANKS = new Set([
  ...VARIADIC, "SMALL", "LARGE", "PERCENTILEEXC", "PERCENTILEINC", "QUARTILEEXC", "QUARTILEINC", "PERCENTRANKEXC",
  "PERCENTRANKINC", "RANKAVG", "RANKEQ", "TRIMMEAN", "MODEMULT", "MODESNGL",
]);
// Functions that build arrays from a size argument, or take a size that could exhaust memory.
const BLOCKED = new Set(["MUNIT", "EXPAND"]);
const ARG_MAX: Record<string, [index: number, max: number]> = {
  BASE: [2, 255], FIXED: [1, 100], DOLLAR: [1, 100], DEC2BIN: [1, 10], DEC2HEX: [1, 10], DEC2OCT: [1, 10],
  BIN2HEX: [1, 10], BIN2OCT: [1, 10], HEX2BIN: [1, 10], HEX2OCT: [1, 10], OCT2BIN: [1, 10], OCT2HEX: [1, 10],
};
const NUMBER = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?$/i;

const err = (error: ErrorCode): ErrorValue => ({ error });
const isErr = (v: unknown): v is ErrorValue => typeof v === "object" && v !== null && "error" in v;

function toNumber(v: Value): number | ErrorValue {
  if (isErr(v)) return v;
  if (typeof v === "number") return v;
  if (v === null || typeof v === "boolean") return Number(v);
  return NUMBER.test(v.trim()) ? Number(v) : err("#VALUE!");
}

function toText(v: Value): string {
  if (v === null) return "";
  if (typeof v === "number") return String(Number(v.toPrecision(15)));
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  return isErr(v) ? v.error : v;
}

function toBool(v: Value): boolean | ErrorValue {
  if (isErr(v)) return v;
  if (typeof v === "string") {
    const upper = v.toUpperCase();
    return upper === "TRUE" || (upper === "FALSE" ? false : err("#VALUE!"));
  }
  return Boolean(v);
}

// Order values as a spreadsheet does: numbers, then text, then booleans.
function compare(a: Value, b: Value): number {
  const blank = (other: Value): Value => (typeof other === "string" ? "" : typeof other === "boolean" ? false : 0);
  const x = a ?? blank(b);
  const y = b ?? blank(a);
  const rank = (v: Value) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  if (rank(x) !== rank(y)) return rank(x) - rank(y);
  const [p, q] = typeof x === "string" ? [x.toLowerCase(), (y as string).toLowerCase()] : [Number(x), Number(y)];
  return p < q ? -1 : p > q ? 1 : 0;
}

// Numbers in ranges count; text, booleans, and blanks in ranges do not.
// A direct argument is coerced, so "5" is 5 and "a" is #VALUE!.
function numbersIn(args: (Value | Matrix)[]): number[] | ErrorValue {
  const out: number[] = [];
  for (const arg of args) {
    if (Array.isArray(arg)) {
      for (const cell of arg.flat()) {
        if (isErr(cell)) return cell;
        if (typeof cell === "number") out.push(cell);
      }
    } else if (arg !== null) {
      const n = toNumber(arg as Value);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}

const NATIVE: Record<string, (args: (Value | Matrix)[]) => Value> = {
  SUM: (a) => { const n = numbersIn(a); return isErr(n) ? n : n.reduce((x, y) => x + y, 0); },
  AVERAGE: (a) => { const n = numbersIn(a); return isErr(n) ? n : n.length === 0 ? err("#DIV/0!") : n.reduce((x, y) => x + y, 0) / n.length; },
  MIN: (a) => { const n = numbersIn(a); return isErr(n) ? n : n.reduce((m, x) => (x < m ? x : m), n[0] ?? 0); },
  MAX: (a) => { const n = numbersIn(a); return isErr(n) ? n : n.reduce((m, x) => (x > m ? x : m), n[0] ?? 0); },
  COUNT: (a) => a.flatMap((x) => (Array.isArray(x) ? x.flat() : [x])).filter((x) => typeof x === "number").length,
  XLOOKUP: ([keyArg, look, ret, missingArg]) => {
    const [key, missing] = [keyArg, missingArg].map((v) => (Array.isArray(v) ? (v[0]?.[0] ?? null) : v)) as [Value, Value | undefined];
    if (isErr(key)) return key;
    if (!Array.isArray(look) || !Array.isArray(ret)) return err("#VALUE!");
    const i = look.flat().findIndex((cell) => cell !== null && compare(cell, key) === 0);
    return i >= 0 ? (ret.flat()[i] ?? null) : (missing ?? err("#N/A"));
  },
  COUNTA: (a) => a.flatMap((x) => (Array.isArray(x) ? x.flat() : [x])).filter((x) => x !== null).length,
};

function toJs(v: Value | Matrix, blank: unknown): unknown {
  if (Array.isArray(v)) return v.map((row) => row.map((cell) => toJs(cell, blank)));
  if (isErr(v)) return new Error(v.error);
  return v ?? blank;
}

const skipBlanks = (v: Value | Matrix): Value | Matrix =>
  Array.isArray(v) ? [v.flat().filter((cell) => typeof cell === "number" || isErr(cell))] : v;

function fromJs(r: unknown): Value {
  if (r instanceof Error) return err(ERRORS.has(r.message) ? (r.message as ErrorCode) : "#VALUE!");
  if (r instanceof Date) {
    const ms = Date.UTC(r.getFullYear(), r.getMonth(), r.getDate(), r.getHours(), r.getMinutes(), r.getSeconds(), r.getMilliseconds());
    return Number.isFinite(ms) ? ms / 86_400_000 + 25_569 : err("#NUM!");
  }
  if (Array.isArray(r)) return fromJs(r.flat(Infinity)[0]);
  if (typeof r === "number") return Number.isFinite(r) ? r : err("#NUM!");
  return typeof r === "string" || typeof r === "boolean" ? r : null;
}

// Computes cell values. Each value() or values() call starts with an
// empty cache, so TODAY and NOW stay fresh and no value goes stale.
export class Engine {
  private readonly cache = new Map<string, Value>();
  // Cells being computed now, with their depth in the chain.
  private readonly visiting = new Map<string, number>();
  private depth = 0;
  // The shallowest open loop. A value computed inside a loop is not cached,
  // so it does not depend on which cell the read started from.
  private taint = Infinity;
  private work = 0;

  constructor(private readonly wb: Workbook) {}

  value(sheet: string, cell: string): Value {
    const { col, row } = parseCell(cell);
    this.reset();
    return this.cellValue(this.wb.sheetInfo(sheet).name, col, row);
  }

  values(sheet: string, range: string): Matrix {
    const name = this.wb.sheetInfo(sheet).name;
    const r = parseRange(range);
    this.reset();
    return Array.from({ length: r.r2 - r.r1 + 1 }, (_, dr) =>
      Array.from({ length: r.c2 - r.c1 + 1 }, (_, dc) => {
        this.work = 0;
        return this.cellValue(name, r.c1 + dc, r.r1 + dr);
      }),
    );
  }

  private reset(): void {
    this.cache.clear();
    this.visiting.clear();
    [this.depth, this.taint, this.work] = [0, Infinity, 0];
  }

  private cellValue(sheet: string, col: number, row: number): Value {
    const a1 = formatCell({ col, row });
    const raw = this.wb.raw(sheet, col, row);
    if (raw.formula === null) return raw.value;
    const key = `${sheet.toLowerCase()}!${a1}`;
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;
    const open = this.visiting.get(key);
    if (open !== undefined) {
      this.taint = Math.min(this.taint, open);
      return err("#CIRCULAR");
    }
    if (this.depth >= MAX_DEPTH) return err("#DEPTH!");
    const mine = this.depth;
    this.visiting.set(key, mine);
    this.depth += 1;
    let out: Value;
    try {
      out = this.scalar(parse(raw.formula, sheet), { home: sheet }) ?? 0;
    } catch (error) {
      if (error instanceof DepthError || error instanceof RangeError) out = err("#DEPTH!");
      else if (error instanceof ParseError) out = err("#VALUE!");
      else throw error;
    } finally {
      this.depth = mine;
      this.visiting.delete(key);
    }
    if (mine < this.taint) this.cache.set(key, out);
    if (mine <= this.taint) this.taint = Infinity;
    return out;
  }

  private matrix(sheet: string, r: RangePos): Matrix {
    return Array.from({ length: r.r2 - r.r1 + 1 }, (_, dr) =>
      Array.from({ length: r.c2 - r.c1 + 1 }, (_, dc) => this.cellValue(sheet, r.c1 + dc, r.r1 + dr)),
    );
  }

  // The sheet a ref points at, or #REF! when it does not exist.
  private target(node: Ref, ctx: Ctx): string | ErrorValue {
    const name = this.wb.sheetNames().find((n) => sameSheet(n, node.sheet ?? ctx.home));
    return name ?? err("#REF!");
  }

  private scalar(node: Node, ctx: Ctx): Value {
    switch (node.t) {
      case "lit":
        return node.v as Value;
      case "ref": {
        const sheet = this.target(node, ctx);
        if (isErr(sheet)) return sheet;
        return node.isRange ? err("#VALUE!") : this.cellValue(sheet, node.range.c1, node.range.r1);
      }
      case "un": {
        const v = this.scalar(node.a, ctx);
        if (node.op === "+" || isErr(v)) return v;
        const n = toNumber(v);
        return isErr(n) ? n : node.op === "-" ? -n : n / 100;
      }
      case "bin":
        return this.binary(node, ctx);
      case "fn":
        return this.call(node, ctx);
    }
  }

  private binary(node: Extract<Node, { t: "bin" }>, ctx: Ctx): Value {
    const a = this.scalar(node.a, ctx);
    const b = this.scalar(node.b, ctx);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    const { op } = node;
    if (op === "&") return toText(a) + toText(b);
    if (PRECEDENCE_COMPARE.has(op)) {
      const c = compare(a, b);
      return { "=": c === 0, "<>": c !== 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0 }[op] as boolean;
    }
    const x = toNumber(a);
    const y = toNumber(b);
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    if (op === "/" && y === 0) return err("#DIV/0!");
    if (op === "^" && x === 0 && y <= 0) return err(y === 0 ? "#NUM!" : "#DIV/0!");
    const r = op === "+" ? x + y : op === "-" ? x - y : op === "*" ? x * y : op === "/" ? x / y : x ** y;
    return Number.isFinite(r) ? r : err("#NUM!");
  }

  // A range is a matrix. A single-cell ref is a matrix too when `asRef` is set, so that
  // functions which skip text and blanks in refs can tell a ref from a typed value.
  private arg(node: Node, ctx: Ctx, asRef: boolean): Value | Matrix {
    if (node.t !== "ref" || !(node.isRange || asRef)) return this.scalar(node, ctx);
    const sheet = this.target(node, ctx);
    if (isErr(sheet)) return sheet;
    const { c1, c2, r1, r2 } = node.range;
    const cells = (c2 - c1 + 1) * (r2 - r1 + 1);
    this.work += cells;
    return cells > MAX_RANGE_CELLS || this.work > MAX_WORK ? err("#VALUE!") : this.matrix(sheet, node.range);
  }

  private call(node: Extract<Node, { t: "fn" }>, ctx: Ctx): Value {
    const { name, args } = node;
    if (name === "IF") {
      if (args.length < 2 || args.length > 3) return err("#VALUE!");
      const test = toBool(this.scalar(args[0], ctx));
      if (isErr(test)) return test;
      return test ? this.scalar(args[1], ctx) : args[2] ? this.scalar(args[2], ctx) : false;
    }
    if (name === "IFS") {
      for (let i = 0; i + 1 < args.length; i += 2) {
        const test = toBool(this.scalar(args[i], ctx));
        if (isErr(test) || test) return isErr(test) ? test : this.scalar(args[i + 1], ctx);
      }
      return args.length % 2 === 0 ? err("#N/A") : err("#VALUE!");
    }
    if (name === "IFERROR") {
      if (args.length !== 2) return err("#VALUE!");
      const v = this.scalar(args[0], ctx);
      return isErr(v) ? this.scalar(args[1], ctx) : v;
    }
    if (BLOCKED.has(name)) return err("#NAME?");
    const values = args.map((a) => this.arg(a, ctx, Object.hasOwn(NATIVE, name) || VARIADIC.has(name)));
    if (Object.hasOwn(NATIVE, name)) return NATIVE[name](values);
    // formulajs AVERAGEIF returns null when the range holds a blank. AVERAGEIFS does not.
    const alias = name === "AVERAGEIF" ? "AVERAGEIFS" : name;
    const fn = Object.hasOwn(formulajs, alias) ? (formulajs as Record<string, unknown>)[alias] : undefined;
    if (typeof fn !== "function") return err("#NAME?");
    const [index, max] = Object.hasOwn(ARG_MAX, name) ? ARG_MAX[name] : [0, Infinity];
    if (Number(values[index]) > max) return err("#NUM!");
    if (name === "REPT" && toText(values[0] as Value).length * Number(values[1]) > MAX_TEXT) return err("#VALUE!");
    const blank = TEXT_FUNCTIONS.has(name) ? "" : BLANK_CELLS.has(name) ? undefined : name === "ISBLANK" ? null : 0;
    const given = name === "AVERAGEIF" ? [values[2] ?? values[0], values[0], values[1]] : SKIP_BLANKS.has(name) ? values.map(skipBlanks) : values;
    try {
      return fromJs(fn(...given.map((v) => toJs(v, blank))));
    } catch {
      return err("#VALUE!");
    }
  }
}

const PRECEDENCE_COMPARE = new Set(["=", "<>", "<", ">", "<=", ">="]);
