import { describe, expect, test } from "vitest";
import { DepthError, ParseError, parse, type Literal, type Node } from "./formula.js";

// Failure modes, listed before the tests:
// - precedence or associativity is wrong (2^3^2 is left-assoc, 1-2-3, -2^2, 1+2*3)
// - "%" and unary signs bind to the wrong operand
// - strings: doubled quotes are kept, an open string is accepted
// - a LOG10( call or a name is read as a cell ref
// - refs lose their sheet, their range, or the isRange flag
// - empty args (SUM(,1)) are dropped or crash
// - stray tokens, unbalanced parens, and bad characters are accepted
// - deep nesting or long chains overflow the stack instead of a DepthError

const lit = (v: Literal): Node => ({ t: "lit", v });
const bin = (op: string, a: Node, b: Node): Node => ({ t: "bin", op, a, b });
const n = (v: number): Node => lit(v);

describe("parse", () => {
  test("accepts a formula with or without the leading =", () => {
    expect(parse("=1", "S")).toEqual(n(1));
    expect(parse("1", "S")).toEqual(n(1));
  });

  test("numbers: decimals, leading dot, exponent", () => {
    expect(parse("1.5", "S")).toEqual(n(1.5));
    expect(parse(".5", "S")).toEqual(n(0.5));
    expect(parse("1e3", "S")).toEqual(n(1000));
    expect(parse("2.5E-1", "S")).toEqual(n(0.25));
  });

  test("precedence and associativity", () => {
    expect(parse("1+2*3", "S")).toEqual(bin("+", n(1), bin("*", n(2), n(3))));
    expect(parse("1-2-3", "S")).toEqual(bin("-", bin("-", n(1), n(2)), n(3)));
    // Like Excel, ^ is left-associative: 2^3^2 is 64.
    expect(parse("2^3^2", "S")).toEqual(bin("^", bin("^", n(2), n(3)), n(2)));
    expect(parse("(1+2)*3", "S")).toEqual(bin("*", bin("+", n(1), n(2)), n(3)));
    expect(parse("1&2=3", "S")).toEqual(bin("=", bin("&", n(1), n(2)), n(3)));
    expect(parse("1+2&3", "S")).toEqual(bin("&", bin("+", n(1), n(2)), n(3)));
  });

  test("unary minus binds tighter than ^, and % is postfix", () => {
    expect(parse("-2^2", "S")).toEqual(bin("^", { t: "un", op: "-", a: n(2) }, n(2)));
    expect(parse("50%", "S")).toEqual({ t: "un", op: "%", a: n(50) });
    expect(parse("50%%", "S")).toEqual({ t: "un", op: "%", a: { t: "un", op: "%", a: n(50) } });
    expect(parse("--1", "S")).toEqual({ t: "un", op: "-", a: { t: "un", op: "-", a: n(1) } });
    expect(parse("+1", "S")).toEqual({ t: "un", op: "+", a: n(1) });
  });

  test("two-character comparison operators", () => {
    for (const op of ["<=", ">=", "<>", "<", ">", "="]) expect(parse(`1${op}2`, "S")).toEqual(bin(op, n(1), n(2)));
  });

  test("strings keep doubled quotes and may hold operators", () => {
    expect(parse('"a""b"', "S")).toEqual(lit('a"b'));
    expect(parse('"1+2"', "S")).toEqual(lit("1+2"));
    expect(parse('""', "S")).toEqual(lit(""));
    expect(parse('"x"&"y"', "S")).toEqual(bin("&", lit("x"), lit("y")));
  });

  test("an unterminated string is a ParseError", () => {
    expect(() => parse('"abc', "S")).toThrow(ParseError);
    expect(() => parse('"abc""', "S")).toThrow(/unterminated/);
  });

  test("booleans, unknown names, and error literals", () => {
    expect(parse("true", "S")).toEqual(lit(true));
    expect(parse("FALSE", "S")).toEqual(lit(false));
    expect(parse("foo", "S")).toEqual(lit({ error: "#NAME?" }));
    for (const e of ["#REF!", "#DIV/0!", "#NAME?", "#VALUE!", "#NUM!", "#N/A", "#CIRCULAR", "#DEPTH!"]) {
      expect(parse(e, "S")).toEqual(lit({ error: e }));
    }
    expect(() => parse("#BOGUS", "S")).toThrow(ParseError);
  });

  test("refs: own sheet, other sheet, quoted sheet, range, absolute", () => {
    expect(parse("A1", "Home")).toEqual({ t: "ref", sheet: null, range: { c1: 1, r1: 1, c2: 1, r2: 1 }, isRange: false });
    expect(parse("$B$2", "Home")).toMatchObject({ t: "ref", sheet: null, range: { c1: 2, r1: 2, c2: 2, r2: 2 } });
    expect(parse("Data!C3", "Home")).toMatchObject({ sheet: "Data", range: { c1: 3, r1: 3 } });
    expect(parse("'My ''Sheet'!A1", "Home")).toMatchObject({ sheet: "My 'Sheet" });
    expect(parse("B2:A1", "Home")).toEqual({ t: "ref", sheet: null, range: { c1: 1, r1: 1, c2: 2, r2: 2 }, isRange: true });
    expect(parse("A1:A1", "Home")).toMatchObject({ isRange: true });
  });

  test("lower-case refs parse as refs", () => {
    expect(parse("a1", "S")).toMatchObject({ t: "ref", range: { c1: 1, r1: 1 } });
  });

  test("functions: names are upper-cased, args nest, empty args are null", () => {
    expect(parse("sum(1,2)", "S")).toEqual({ t: "fn", name: "SUM", args: [n(1), n(2)] });
    expect(parse("NOW()", "S")).toEqual({ t: "fn", name: "NOW", args: [] });
    expect(parse("IF(,1,)", "S")).toEqual({ t: "fn", name: "IF", args: [lit(null), n(1), lit(null)] });
    expect(parse("SUM(A1:B2, MAX(1,2))", "S")).toMatchObject({ args: [{ t: "ref", isRange: true }, { t: "fn", name: "MAX" }] });
  });

  test("a function name that looks like a cell is still a function", () => {
    expect(parse("LOG10(100)", "S")).toEqual({ t: "fn", name: "LOG10", args: [n(100)] });
    expect(parse("ATAN2(1,1)", "S")).toMatchObject({ t: "fn", name: "ATAN2" });
  });

  test("dotted function names", () => {
    expect(parse("STDEV.S(1,2)", "S")).toMatchObject({ t: "fn", name: "STDEV.S" });
  });

  test("whitespace is ignored", () => {
    expect(parse("  1 +\n 2\t", "S")).toEqual(bin("+", n(1), n(2)));
  });

  test.each(["", "=", "1+", "*1", "(1", "1)", "SUM(1", "SUM(1,", "1 2", "1,2", "@", "A1 B1", "SUM(1))", "()", "1+*2"])(
    "%j is a ParseError",
    (text) => {
      expect(() => parse(text, "S")).toThrow(ParseError);
    },
  );

  test("out-of-range refs are not refs", () => {
    expect(() => parse("XFE1", "S")).not.toThrow();
    expect(parse("XFE1", "S")).toEqual(lit({ error: "#NAME?" }));
    expect(parse("A1048577", "S")).toEqual(lit({ error: "#NAME?" }));
    expect(parse("A0", "S")).toEqual(lit({ error: "#NAME?" }));
  });

  test("deep nesting throws DepthError, not a stack overflow", () => {
    expect(() => parse(`${"(".repeat(5000)}1${")".repeat(5000)}`, "S")).toThrow(DepthError);
    expect(() => parse("-".repeat(5000) + "1", "S")).toThrow(DepthError);
    expect(() => parse(`${"SUM(".repeat(5000)}1${")".repeat(5000)}`, "S")).toThrow(DepthError);
    expect(() => parse(`${"1^".repeat(5000)}1`, "S")).toThrow(DepthError);
    expect(new DepthError("x")).toBeInstanceOf(ParseError);
  });

  test("long operator chains throw DepthError; moderate chains parse", () => {
    expect(() => parse(`1${"+1".repeat(1500)}`, "S")).toThrow(DepthError);
    expect(() => parse(`1${"+1".repeat(500)}`, "S")).not.toThrow();
  });

  test("the limits sit at 99/100 nested parens and 1000/1001 chained operators", () => {
    expect(() => parse(`${"(".repeat(99)}1${")".repeat(99)}`, "S")).not.toThrow();
    expect(() => parse(`${"(".repeat(100)}1${")".repeat(100)}`, "S")).toThrow(DepthError);
    expect(() => parse(`1${"+1".repeat(1000)}`, "S")).not.toThrow();
    expect(() => parse(`1${"+1".repeat(1001)}`, "S")).toThrow(DepthError);
  });

  test("a colon in a quoted sheet name does not make a single cell a range", () => {
    expect(parse("'a:b'!A1", "S")).toMatchObject({ sheet: "a:b", isRange: false });
    expect(parse("'a:b'!A1:B2", "S")).toMatchObject({ sheet: "a:b", isRange: true });
  });
});
