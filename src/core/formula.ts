import type { RangePos } from "./a1.js";
import { readRef } from "./refs.js";

export type Literal = number | string | boolean | null | { error: string };

export type Node =
  | { t: "lit"; v: Literal }
  | { t: "ref"; sheet: string | null; range: RangePos; isRange: boolean }
  | { t: "fn"; name: string; args: Node[] }
  | { t: "un"; op: "-" | "+" | "%"; a: Node }
  | { t: "bin"; op: string; a: Node; b: Node };

export class ParseError extends Error {}
// A formula nested or chained too deep to evaluate safely.
export class DepthError extends ParseError {}

const MAX_NESTING = 100;
const MAX_CHAIN = 1_000;

type Token = { k: "num" | "str" | "err" | "ref" | "fn" | "name" | "op"; v: string; ref?: ReturnType<typeof readRef> };

const NUMBER = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
const ERROR = /#(?:REF!|DIV\/0!|NAME\?|VALUE!|NUM!|N\/A|CIRCULAR|DEPTH!)/y;
const IDENT = /[A-Za-z_][\w.]*/y;
const OPERATOR = /<=|>=|<>|[-+*/^&=<>(),%]/y;
const PRECEDENCE: Record<string, number> = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5 };

function sticky(re: RegExp, text: string, at: number): string | null {
  re.lastIndex = at;
  return re.exec(text)?.[0] ?? null;
}

function tokenize(text: string, home: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  while (i < text.length) {
    if (/\s/.test(text[i])) {
      i += 1;
      continue;
    }
    if (text[i] === '"') {
      let j = i + 1;
      while (j < text.length && (text[j] !== '"' || text[j + 1] === '"')) j += text[j] === '"' ? 2 : 1;
      if (j >= text.length) throw new ParseError("unterminated string");
      tokens.push({ k: "str", v: text.slice(i + 1, j).replaceAll('""', '"') });
      i = j + 1;
      continue;
    }
    const raw = sticky(NUMBER, text, i);
    const ref = readRef(text, i, home);
    let token: Token;
    if (raw !== null) token = { k: "num", v: raw };
    else if (ref !== null) token = { k: "ref", v: ref.text, ref };
    else if (text[i] === "#") token = { k: "err", v: sticky(ERROR, text, i) ?? "" };
    else if (sticky(IDENT, text, i) !== null) {
      const word = sticky(IDENT, text, i) as string;
      token = { k: text[i + word.length] === "(" ? "fn" : "name", v: word };
    } else token = { k: "op", v: sticky(OPERATOR, text, i) ?? "" };
    if (token.v === "") throw new ParseError(`unexpected character "${text[i]}"`);
    tokens.push(token);
    i += token.v.length;
  }
  return tokens;
}

// Parse a formula, with or without its leading "=". Precedence follows
// the usual spreadsheet order: unary minus binds tighter than ^.
export function parse(formula: string, home: string): Node {
  const tokens = tokenize(formula.replace(/^=/, ""), home);
  let at = 0;
  const peek = () => tokens[at];
  const take = (op?: string): Token => {
    const token = tokens[at];
    if (token === undefined || (op !== undefined && token.v !== op)) throw new ParseError(`expected ${op ?? "more input"}`);
    at += 1;
    return token;
  };
  const isOp = (v: string) => peek()?.k === "op" && peek().v === v;

  let level = 0;
  const nested = <T>(run: () => T): T => {
    if (++level > MAX_NESTING) throw new DepthError("formula nests too deep");
    try {
      return run();
    } finally {
      level -= 1;
    }
  };
  const binary = (min: number): Node =>
    nested(() => {
      let a = unary();
      let chain = 0;
      while (peek()?.k === "op" && (PRECEDENCE[peek().v] ?? 0) >= min) {
        if (++chain > MAX_CHAIN) throw new DepthError("formula chains too many operators");
        const op = take().v;
        a = { t: "bin", op, a, b: binary(PRECEDENCE[op] + 1) };
      }
      return a;
    });
  const unary = (): Node => {
    if (isOp("-") || isOp("+")) return nested(() => ({ t: "un", op: take().v as "-" | "+", a: unary() }));
    let a = primary();
    while (isOp("%")) {
      take();
      a = { t: "un", op: "%", a };
    }
    return a;
  };
  const primary = (): Node => {
    const token = take();
    switch (token.k) {
      case "num":
        return { t: "lit", v: Number(token.v) };
      case "str":
        return { t: "lit", v: token.v };
      case "err":
        return { t: "lit", v: { error: token.v } };
      case "ref": {
        const { ref } = token.ref!;
        return { t: "ref", sheet: ref.sheet, range: ref.range, isRange: token.ref!.flags.isRange };
      }
      case "name": {
        const word = token.v.toUpperCase();
        return { t: "lit", v: word === "TRUE" || word === "FALSE" ? word === "TRUE" : { error: "#NAME?" } };
      }
      case "fn": {
        take("(");
        const args: Node[] = [];
        if (!isOp(")")) {
          do args.push(isOp(",") || isOp(")") ? { t: "lit", v: null } : binary(1));
          while (isOp(",") && take());
        }
        take(")");
        return { t: "fn", name: token.v.toUpperCase(), args };
      }
      default:
        if (token.v !== "(") throw new ParseError(`unexpected ${token.v}`);
        {
          const inner = binary(1);
          take(")");
          return inner;
        }
    }
  };
  const root = binary(1);
  if (at < tokens.length) throw new ParseError(`unexpected ${tokens[at].v}`);
  return root;
}
