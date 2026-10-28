import * as formulajs from "@formulajs/formulajs";
import { formatCell, parseCell, parseRange, type RangePos } from "./a1.js";
import { DepthError, parse, ParseError, type Node } from "./formula.js";
import { sameSheet } from "./refs.js";
import { MAX_TEXT, type Workbook } from "./workbook.js";

export type ErrorCode = "#REF!" | "#DIV/0!" | "#NAME?" | "#VALUE!" | "#NUM!" | "#N/A" | "#CIRCULAR" | "#DEPTH!";
export type ErrorValue = { error: ErrorCode };
export type Value = number | string | boolean | null | ErrorValue;
type Matrix = Value[][];
type Pos = { sheet: string; col: number; row: number; key: string };
// The state of one formula while it is evaluated. `circ` holds cells met while they were
// open: they read as #CIRCULAR. `cyc` is set when the cell lies on a loop.
interface Frame {
  pos: Pos;
  formula: string;
  circ: Set<string>;
  cyc: boolean;
  points: Set<string>;
  ranges: { sheet: string; r: RangePos }[];
  volatile: boolean;
}
type Ctx = { home: string; frame: Frame };
// Thrown when a formula reads a formula cell that has no value yet. The driver computes those
// cells first, then evaluates the formula again.
class Need {
  constructor(readonly cells: Pos[]) {}
}
const VOLATILE = new Set(["NOW", "TODAY", "RAND", "RANDBETWEEN"]);
type Ref = Extract<Node, { t: "ref" }>;

const ERRORS = new Set<string>(["#REF!", "#DIV/0!", "#NAME?", "#VALUE!", "#NUM!", "#N/A", "#CIRCULAR", "#DEPTH!"]);
const MAX_RANGE_CELLS = 1_000_000;
// Range cells one formula may read.
const MAX_WORK = 10_000_000;
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

const keyOf = (sheet: string, address: string) => `${sheet.toLowerCase()}!${address}`;

// Computes cell values and keeps them. A value stays until a change to the workbook reaches it.
// The engine records what each formula read, including the cells and ranges in the branch of an
// IF that ran. A change to a cell clears only the values that read it, directly or through
// other cells. Evaluation uses an explicit stack, so a chain of any length cannot overflow.
// Call dispose() when you are done with an engine whose workbook lives on.
export class Engine {
  private readonly cache = new Map<string, Value>();
  private readonly asts = new Map<string, Node>();
  // What each cached formula read, and the reverse: which formulas read each cell.
  private readonly reads = new Map<string, Set<string>>();
  private readonly readers = new Map<string, Set<string>>();
  private readonly rangesOf = new Map<string, { sheet: string; r: RangePos }[]>();
  private readonly rangeReaders = new Map<string, Set<string>>();
  private readonly volatile = new Set<string>();
  private names: string[] | null = null;
  private work = 0;
  private readonly stop: () => void;
  // How many formulas this engine has computed. A read that hits the cache does not add to it.
  evaluations = 0;

  constructor(private readonly wb: Workbook) {
    this.stop = wb.onChange((change) => {
      if (change.all) return this.clear();
      this.invalidate(change.cells.map(({ sheet, address }) => keyOf(sheet, address)), true);
    });
  }

  dispose(): void {
    this.stop();
  }

  value(sheet: string, cell: string): Value {
    const { col, row } = parseCell(cell);
    this.refresh();
    return this.resolve(this.wb.sheetInfo(sheet).name, col, row);
  }

  values(sheet: string, range: string): Matrix {
    const name = this.wb.sheetInfo(sheet).name;
    const r = parseRange(range);
    this.refresh();
    return Array.from({ length: r.r2 - r.r1 + 1 }, (_, dr) =>
      Array.from({ length: r.c2 - r.c1 + 1 }, (_, dc) => this.resolve(name, r.c1 + dc, r.r1 + dr)),
    );
  }

  private clear(): void {
    for (const map of [this.cache, this.asts, this.reads, this.readers, this.rangesOf, this.rangeReaders]) map.clear();
    this.volatile.clear();
    this.names = null;
  }

  // NOW, TODAY, and RAND are new on every read, and so are the cells that read them.
  private refresh(): void {
    if (this.volatile.size > 0) this.invalidate([...this.volatile], false);
  }

  // Clear the values of `seeds` and of every cell that reads one of them. A seed whose content
  // changed also loses its parsed formula.
  private invalidate(seeds: string[], contentChanged: boolean): void {
    const seen = new Set(seeds);
    const queue = [...seeds];
    if (contentChanged) for (const seed of seeds) this.asts.delete(seed);
    while (queue.length > 0) {
      const key = queue.pop() as string;
      for (const reader of this.readersOf(key)) if (!seen.has(reader)) (seen.add(reader), queue.push(reader));
      this.drop(key);
    }
  }

  private readersOf(key: string): string[] {
    const out = [...(this.readers.get(key) ?? [])];
    const cut = key.lastIndexOf("!");
    const sheet = key.slice(0, cut);
    const at = parseCell(key.slice(cut + 1));
    for (const owner of this.rangeReaders.get(sheet) ?? []) {
      const hit = this.rangesOf.get(owner)?.some((x) => x.sheet === sheet && at.col >= x.r.c1 && at.col <= x.r.c2 && at.row >= x.r.r1 && at.row <= x.r.r2);
      if (hit) out.push(owner);
    }
    return out;
  }

  private drop(key: string): void {
    this.cache.delete(key);
    this.volatile.delete(key);
    for (const read of this.reads.get(key) ?? []) {
      const set = this.readers.get(read);
      if (set?.delete(key) && set.size === 0) this.readers.delete(read);
    }
    this.reads.delete(key);
    for (const { sheet } of this.rangesOf.get(key) ?? []) this.rangeReaders.get(sheet)?.delete(key);
    this.rangesOf.delete(key);
  }

  private resolve(sheet: string, col: number, row: number): Value {
    const raw = this.wb.raw(sheet, col, row);
    if (raw.formula === null) return raw.value;
    const pos = { sheet, col, row, key: keyOf(sheet, formatCell({ col, row })) };
    if (!this.cache.has(pos.key)) this.run(pos);
    return this.cache.get(pos.key) as Value;
  }

  private frame(pos: Pos): Frame {
    const formula = this.wb.raw(pos.sheet, pos.col, pos.row).formula as string;
    return { pos, formula, circ: new Set(), cyc: false, points: new Set(), ranges: [], volatile: false };
  }

  // Compute `root` and the formula cells it needs, nearest first, on an explicit stack. A formula
  // that needs a cell with no value yet waits, and the driver runs it again once the cell is ready.
  // A cell it needs that is already waiting is a loop: the cells from there to here are on it.
  private run(root: Pos): void {
    const stack: Frame[] = [this.frame(root)];
    const waiting = new Map<string, number>();
    const chain: Frame[] = [];
    while (stack.length > 0) {
      const f = stack[stack.length - 1];
      const { key } = f.pos;
      if (this.cache.has(key)) {
        stack.pop();
        continue;
      }
      let out: Value;
      try {
        out = this.attempt(f);
      } catch (need) {
        if (!(need instanceof Need)) throw need;
        if (!waiting.has(key)) (waiting.set(key, chain.length), chain.push(f));
        for (const m of need.cells) {
          const at = waiting.get(m.key);
          if (at !== undefined) {
            f.circ.add(m.key);
            for (let i = at; i < chain.length; i += 1) chain[i].cyc = true;
          } else if (!this.cache.has(m.key)) stack.push(this.frame(m));
        }
        continue;
      }
      this.finish(f, f.cyc ? err("#CIRCULAR") : out);
      if (waiting.delete(key)) chain.pop();
      stack.pop();
    }
  }

  private attempt(f: Frame): Value {
    f.points.clear();
    f.ranges.length = 0;
    f.volatile = false;
    this.work = 0;
    try {
      let ast = this.asts.get(f.pos.key);
      if (ast === undefined) (ast = parse(f.formula, f.pos.sheet), this.asts.set(f.pos.key, ast));
      return this.scalar(ast, { home: f.pos.sheet, frame: f }) ?? 0;
    } catch (error) {
      if (error instanceof Need) throw error;
      if (error instanceof DepthError || error instanceof RangeError) return err("#DEPTH!");
      if (error instanceof ParseError) return err("#VALUE!");
      throw error;
    }
  }

  private finish(f: Frame, out: Value): void {
    const { key } = f.pos;
    this.evaluations += 1;
    this.cache.set(key, out);
    this.reads.set(key, f.points);
    for (const read of f.points) {
      const set = this.readers.get(read) ?? new Set<string>();
      this.readers.set(read, set.add(key));
    }
    if (f.ranges.length > 0) {
      this.rangesOf.set(key, f.ranges);
      for (const { sheet } of f.ranges) this.rangeReaders.set(sheet, (this.rangeReaders.get(sheet) ?? new Set<string>()).add(key));
    }
    if (f.volatile) this.volatile.add(key);
  }

  // The value of a cell, recorded as something this formula read. A formula cell with no value
  // yet is a Need, or is added to `missing` when the caller reads a whole range.
  private read(f: Frame, sheet: string, col: number, row: number, missing?: Pos[]): Value {
    const raw = this.wb.raw(sheet, col, row);
    const key = keyOf(sheet, formatCell({ col, row }));
    if (missing === undefined) f.points.add(key);
    if (raw.formula === null) return raw.value;
    if (f.circ.has(key)) return err("#CIRCULAR");
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;
    if (missing === undefined) throw new Need([{ sheet, col, row, key }]);
    missing.push({ sheet, col, row, key });
    return null;
  }

  private matrix(sheet: string, r: RangePos, f: Frame): Matrix {
    f.ranges.push({ sheet: sheet.toLowerCase(), r });
    const missing: Pos[] = [];
    const rows = Array.from({ length: r.r2 - r.r1 + 1 }, (_, dr) =>
      Array.from({ length: r.c2 - r.c1 + 1 }, (_, dc) => this.read(f, sheet, r.c1 + dc, r.r1 + dr, missing)),
    );
    if (missing.length > 0) throw new Need(missing);
    return rows;
  }

  // The sheet a ref points at, or #REF! when it does not exist.
  private target(node: Ref, ctx: Ctx): string | ErrorValue {
    this.names ??= this.wb.sheetNames();
    const name = this.names.find((n) => sameSheet(n, node.sheet ?? ctx.home));
    return name ?? err("#REF!");
  }

  private scalar(node: Node, ctx: Ctx): Value {
    switch (node.t) {
      case "lit":
        return node.v as Value;
      case "ref": {
        const sheet = this.target(node, ctx);
        if (isErr(sheet)) return sheet;
        return node.isRange ? err("#VALUE!") : this.read(ctx.frame, sheet, node.range.c1, node.range.r1);
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
    return cells > MAX_RANGE_CELLS || this.work > MAX_WORK ? err("#VALUE!") : this.matrix(sheet, node.range, ctx.frame);
  }

  private call(node: Extract<Node, { t: "fn" }>, ctx: Ctx): Value {
    const { name, args } = node;
    if (VOLATILE.has(name)) ctx.frame.volatile = true;
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
