import { describe, expect, test } from "vitest";
import { Engine, type ErrorCode, type Value } from "./engine.js";
import { Workbook, type CellInput } from "./workbook.js";

// Failure modes, listed before the code:
// - a self ref or a loop hangs or returns a number; a ref to a deleted row or sheet is not #REF!
// - 1/0, unknown names, text in math, a range in scalar context, a half-typed formula, and an
//   error in a precedent do not give #DIV/0!, #NAME?, #VALUE!, #VALUE!, #VALUE!, and the same error
// - IF or IFERROR runs the branch it must skip, so 1/0 leaks out
// - empty cells break SUM, COUNT, MAX, AVERAGE; "5"+1, "a"+1, TRUE+1, ""+1, "a"="A", 1="1" coerce
//   wrong; SUM counts a text cell in a range
// - 1+2*3, -2^2, 2^3^2, 10-4-3, 50%, "a"&1+2, 2*-3 come out wrong
// - a read after an edit, a row insert, a delete, or a rename shows a stale value
// - a blank cell counts as 0 in PRODUCT, MEDIAN, SMALL, LARGE, COUNTIF, SUMIF, MATCH, ISBLANK, or as 0
//   instead of "" in CONCATENATE, SUBSTITUTE, FIND, EXACT; text or TRUE from a single-cell ref counts in SUM
// - MIN or MAX over 200,000 numbers throws RangeError; 0^0 and 0^-1 give 1 and #NUM!
// - REPT, BASE, DEC2BIN, MUNIT, or EXPAND can build a huge string or array
// - a 5,000-deep chain, 4,000 parens, 5,000 minus signs, 2,000 nested IFs, or a 3,000-term sum throws
//   RangeError and fails every read that covers the cell, instead of one cell showing #DEPTH!
// - a failed deep read leaves a stale key, so the next read shows #CIRCULAR for a cell with no loop
// - A1 = IFERROR(B1, 5) with B1 = A1 reads as 5 through value() but #CIRCULAR through values()

const me = { id: "u", kind: "human" } as const;

function setup(cells: Record<string, CellInput> = {}) {
  const wb = new Workbook();
  wb.addSheet("Sheet1");
  for (const [at, v] of Object.entries(cells)) wb.writeRange("Sheet1", at, [[v]], me);
  return { wb, engine: new Engine(wb) };
}
const calc = (formula: string, cells: Record<string, CellInput> = {}): Value => {
  const { engine } = setup({ ...cells, Z99: formula });
  return engine.value("Sheet1", "Z99");
};
const err = (error: ErrorCode): Value => ({ error });

const fixture = { A1: "pear", B1: 3, A2: "apple", B2: 5, A3: "fig", B3: 7, C1: " Hi  there ", A5: 5, A6: "text", A7: true, F1: 2, F3: 3, F4: "x", F5: true };
const cases: [string, Value][] = [
  // errors
  ["=1/0", err("#DIV/0!")], ["=FOO(1)", err("#NAME?")], ["=foo", err("#NAME?")],
  ['="a"+1', err("#VALUE!")], ['=""+1', err("#VALUE!")], ["=D1:D3+1", err("#VALUE!")],
  ["=Nope!A1", err("#REF!")], ["=#REF!+1", err("#REF!")], ["=1+", err("#VALUE!")],
  ["=(1", err("#VALUE!")], ["=SUM(", err("#VALUE!")], ["=SUM(B9)+1/0", err("#DIV/0!")],
  ["=IF(TRUE,1,1/0)", 1], ["=IF(FALSE,1/0,2)", 2], ["=IFERROR(1/0,7)", 7], ["=IFERROR(5,1/0)", 5],
  ["=IFS(FALSE,1,TRUE,2)", 2], ["=IFS(FALSE,1)", err("#N/A")],
  // empty cells and coercion
  ["=SUM(D1:D3)", 0], ["=COUNT(D1:D3)", 0], ["=MAX(D1:D3)", 0], ["=AVERAGE(D1:D3)", err("#DIV/0!")],
  ["=D1+1", 1], ['=D1&"x"', "x"], ['="5"+1', 6], ["=TRUE+1", 2], ['="a"="A"', true], ['=1="1"', false],
  ["=1<\"a\"", true], ["=SUM(1,2,3)", 6], ["=SUM(A5:A6)", 5], ["=COUNTA(A5:A7)", 3], ["=COUNT(A5:A7)", 1],
  // precedence
  ["=1+2*3", 7], ["=(1+2)*3", 9], ["=-2^2", 4], ["=2^3^2", 64], ["=10-4-3", 3], ["=2*3%", 0.06],
  ['="a"&1+2', "a3"], ["=2*-3", -6], ["=1+1=2", true], ["=0.1+0.2", 0.30000000000000004], ['=0.1+0.2&""', "0.3"],
  // functions
  ['=VLOOKUP("apple",A1:B3,2,FALSE)', 5], ['=VLOOKUP("kiwi",A1:B3,2,FALSE)', err("#N/A")],
  ['=XLOOKUP("fig",A1:A3,B1:B3)', 7], ['=XLOOKUP("kiwi",A1:A3,B1:B3,0)', 0], ["=INDEX(A1:B3,3,2)", 7],
  ['=MATCH("fig",A1:A3,0)', 3], ['=SUMIF(B1:B3,">3")', 12], ['=SUMIFS(B1:B3,B1:B3,">3",A1:A3,"fig")', 7],
  ['=COUNTIF(B1:B3,">=5")', 2], ['=COUNTIFS(B1:B3,">=5",A1:A3,"fig")', 1], ["=AVERAGE(B1:B3)", 5],
  ["=MIN(B1:B3)", 3], ["=ROUND(2.567,2)", 2.57], ["=ROUNDUP(2.561,2)", 2.57], ["=ROUNDDOWN(-2.569,2)", -2.56],
  ["=ABS(-4)", 4], ["=AND(TRUE,B1>1)", true], ["=OR(FALSE,B1>5)", false], ["=NOT(FALSE)", true],
  ['=CONCAT(A1,"-",B1)', "pear-3"], ["=LEFT(A1,2)&RIGHT(A1,2)&MID(A1,2,2)", "pearea"], ["=LEN(A1)", 4],
  ['=UPPER(A1)&LOWER("AB")', "PEARab"], ["=TRIM(C1)", "Hi there"], ['=TEXT(1234.5,"0.00")', "1234.50"],
  ["=DATE(2024,1,15)", 45306], ["=DATE(2024,1,15)-DATE(2024,1,1)", 14], ["=TODAY()>45000", true],
  ["=NOW()>45000", true], ["=SUM(Data!A1:A2)", 3],
  // blanks in ranges are skipped, not zeros
  ["=PRODUCT(F1:F3)", 6], ["=PRODUCT(F1:F5)", 6], ["=MEDIAN(F1:F3)", 2.5], ["=SMALL(F1:F3,1)", 2], ["=LARGE(F1:F3,2)", 2],
  ["=ROUND(STDEVS(F1:F3),4)", 0.7071], ['=COUNTIF(F1:F3,"<5")', 2], ['=SUMIF(F1:F3,"<5")', 5],
  ['=AVERAGEIF(F1:F3,"<5")', 2.5], ["=COUNTBLANK(F1:F3)", 1], ["=ISBLANK(F2)", true], ["=ISBLANK(F1)", false],
  ["=MATCH(0,F1:F3,0)", err("#N/A")], ['=CONCATENATE("a",F2,"b")', "ab"], ['=SUBSTITUTE(F2,"a","b")', ""],
  ['=FIND("a",F2)', err("#VALUE!")], ['=EXACT(F2,"")', true], ["=ROUND(F2,1)", 0],
  // a single-cell ref is a reference: text, TRUE, and blanks in it do not count
  ["=SUM(F4)", 0], ["=SUM(F5)", 0], ["=SUM(F2)", 0], ["=SUM(F1,F4,F3)", 5], ["=AVERAGE(F4)", err("#DIV/0!")],
  ["=COUNT(F4)", 0], ["=COUNT(F1)", 1], ["=MAX(F4)", 0], ['=SUM("5")', 5], ["=PRODUCT(F1,F4,F3)", 6],
  // edge cases and limits
  ["=0^0", err("#NUM!")], ["=0^-1", err("#DIV/0!")], ['=REPT("ab",3)', "ababab"], ['=REPT("ab",1000000)', err("#VALUE!")],
  ["=BASE(5,2,1000000000)", err("#NUM!")], ["=DEC2BIN(5,1000000)", err("#NUM!")], ["=MUNIT(5)", err("#NAME?")],
  ["=EXPAND(F1:F3,1000000,1000000)", err("#NAME?")],
  // deep input gives one cell error
  [`=${"(".repeat(4000)}1${")".repeat(4000)}`, err("#DEPTH!")], [`=${"-".repeat(5000)}1`, err("#DEPTH!")],
  [`=${"IF(1,".repeat(1000)}1${")".repeat(1000)}`, err("#DEPTH!")], [`=1${"+1".repeat(3000)}`, err("#DEPTH!")],
  [`=1${"+1".repeat(900)}`, 901],
];

describe("formulas", () => {
  const titled = cases.map(([formula, expected]) => [formula.length > 60 ? `${formula.slice(0, 24)}... (${formula.length} chars)` : formula, formula, expected] as const);
  test.each(titled)("%s", (_title, formula, expected) => {
    const { wb, engine } = setup({ ...fixture, Z99: formula });
    wb.addSheet("Data");
    wb.writeRange("Data", "A1", [[1], [2]], me);
    expect(engine.value("Sheet1", "Z99")).toEqual(expected);
  });

  test("an error reaches its dependents", () => {
    expect(calc("=B1+1", { B1: "=1/0" })).toEqual(err("#DIV/0!"));
    expect(calc("=SUM(B1:B2)", { B1: "=1/0", B2: 1 })).toEqual(err("#DIV/0!"));
  });

  test("a self ref and a loop give #CIRCULAR, and the loop clears after a fix", () => {
    const { wb, engine } = setup({ A1: "=A1+1", B1: "=C1", C1: "=B1", D1: "=C1+1" });
    expect(engine.values("Sheet1", "A1:D1")[0].map((v) => (v as { error: string }).error)).toEqual(Array(4).fill("#CIRCULAR"));
    wb.writeRange("Sheet1", "C1", [[5]], me);
    expect(engine.values("Sheet1", "B1:D1")).toEqual([[5, 5, 6]]);
  });

  test("deleted rows and deleted sheets give #REF!", () => {
    const { wb, engine } = setup({ A1: 1, A2: 2, B1: "=A2*2", B3: "=SUM(A1:A2)" });
    wb.addSheet("Data");
    wb.writeRange("Sheet1", "C1", [["=Data!A1"]], me);
    wb.deleteRows("Sheet1", 2, 1);
    expect(engine.value("Sheet1", "B1")).toEqual(err("#REF!"));
    expect(engine.value("Sheet1", "B2")).toBe(1);
    wb.deleteSheet("Data");
    expect(engine.value("Sheet1", "C1")).toEqual(err("#REF!"));
  });
});

describe("after edits", () => {
  test("reads see edits, inserts, and renames; TODAY and NOW stay live", async () => {
    const { wb, engine } = setup({ A1: 1, A2: 2, B1: "=SUM(A1:A2)", C1: "=NOW()" });
    wb.addSheet("Two");
    wb.writeRange("Two", "A1", [["=Sheet1!B1"]], me);
    expect(engine.value("Two", "A1")).toBe(3);
    wb.insertRows("Sheet1", 2, 1);
    wb.writeRange("Sheet1", "A2", [[10]], me);
    expect(engine.value("Two", "A1")).toBe(13);
    wb.renameSheet("Sheet1", "Main");
    expect(engine.value("Two", "A1")).toBe(13);
    const first = engine.value("Main", "C1") as number;
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(engine.value("Main", "C1") as number).toBeGreaterThan(first);
  });
});

describe("large and deep input", () => {
  const chain = (n: number) => {
    const { wb, engine } = setup();
    wb.writeRange("Sheet1", "A1", [[1]], me);
    wb.writeRange("Sheet1", "B1", Array.from({ length: n }, (_, i) => [i === 0 ? "=A1" : `=B${i}+1`]), me);
    return { wb, engine };
  };

  test("MIN and MAX read 200,000 numbers", { timeout: 60_000 }, () => {
    const { wb, engine } = setup();
    wb.writeRange("Sheet1", "A1", Array.from({ length: 200_000 }, (_, i) => [i + 1]), me);
    wb.writeRange("Sheet1", "B1", [["=MAX(A1:A200000)", "=MIN(A1:A200000)", "=SUM(A1:A200000)"]], me);
    expect(engine.values("Sheet1", "B1:D1")).toEqual([[200_000, 1, 20_000_100_000]]);
  });

  test("a deep chain is a cell error, never a throw, and the next read has no stale state", () => {
    const { wb, engine } = chain(5_000);
    const last = engine.value("Sheet1", "B5000");
    expect(last).toEqual(err("#DEPTH!"));
    expect(engine.value("Sheet1", "B3")).toBe(3);
    expect(engine.values("Sheet1", "A1:B3")).toEqual([[1, 1], [null, 2], [null, 3]]);
    wb.writeRange("Sheet1", "D1", [["=SUM(A1:B2)"]], me);
    expect(engine.value("Sheet1", "D1")).toBe(4);
  });

  test("reading the chain from the top works and the poison cell stays in its own cell", () => {
    const { engine } = chain(5_000);
    const rows = engine.values("Sheet1", "B1:B5000");
    expect(rows[4_999]).toEqual([5_000]);
    const poisoned = setup({ A1: 1, B1: `=${"(".repeat(4000)}1${")".repeat(4000)}`, C1: 3 });
    expect(poisoned.engine.values("Sheet1", "A1:C1")).toEqual([[1, err("#DEPTH!"), 3]]);
  });

  test("a loop with IFERROR reads the same through value() and values()", () => {
    const cells = { A1: "=IFERROR(B1,5)", B1: "=A1" };
    const single = [setup(cells).engine.value("Sheet1", "A1"), setup(cells).engine.value("Sheet1", "B1")];
    expect(setup(cells).engine.values("Sheet1", "A1:B1")).toEqual([single]);
  });
});
