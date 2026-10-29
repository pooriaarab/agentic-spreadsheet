import { describe, expect, test } from "vitest";
import { Workbook, type Principal } from "./workbook.js";

// Failure modes, listed before the code. Rows and columns share one
// code path, so rows get the full list.
// - a range that ends just above an insert grows, or one that spans it does not
// - refs in a string literal, in LOG10(, or on another sheet get rewritten
// - refs from other sheets to this sheet are missed, or quoted names break
// - a deleted cell, or a fully deleted range, does not become #REF!
// - a partly deleted range does not shrink
// - a moved cell loses its author, or a deleted cell leaves a key
// - widths and frozen rows stay put when rows or columns move
// - a rename leaves stale refs, a delete leaves live refs, the last sheet goes
// - a bad position or count passes silently, or a count of 2,000,000 passes
// - an insert pushes a cell or a column width off the grid and drops it, or leaves the edit half done
// - a ref that an insert pushes past the last row stays as text that no longer parses as a ref
// Edits made at the same time by two writers are in interleave.test.ts.

const human: Principal = { id: "u1", kind: "human" };
const agent: Principal = { id: "a1", kind: "agent_of_human" };

function book(): Workbook {
  const wb = new Workbook();
  wb.addSheet("Sheet1");
  return wb;
}
const formulas = (wb: Workbook, sheet: string, range: string) =>
  wb.readRange(sheet, range).flat().map((c) => c.formula ?? c.value);

describe("insert and delete rows", () => {
  test("cells below shift, cells above stay, authors move with the cell", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[1], [2], [3]], agent);
    wb.insertRows("Sheet1", 2, 2);
    expect(wb.readRange("Sheet1", "A1:A5").flat().map((c) => c.value)).toEqual([1, null, null, 2, 3]);
    expect(wb.readRange("Sheet1", "A5")[0][0].author?.machine).toBe(true);
  });

  test("ranges grow when an insert lands inside, and hold when it lands just below", () => {
    const wb = book();
    wb.writeRange("Sheet1", "B1", [["=SUM(A1:A3)", "=SUM(A1:A3)", "=SUM($A$1:A$3)"]], human);
    wb.insertRows("Sheet1", 2, 1);
    expect(formulas(wb, "Sheet1", "B1:D1")).toEqual(["=SUM(A1:A4)", "=SUM(A1:A4)", "=SUM($A$1:A$4)"]);
    wb.insertRows("Sheet1", 5, 1);
    expect(formulas(wb, "Sheet1", "B1:B1")).toEqual(["=SUM(A1:A4)"]);
  });

  test("refs to moved cells follow; literals, function names, and other sheets stay", () => {
    const wb = book();
    wb.addSheet("Data");
    wb.writeRange("Sheet1", "C1", [['=A5+Data!A5+LOG10(A5)&"A5"']], human);
    wb.insertRows("Sheet1", 1, 1);
    expect(formulas(wb, "Sheet1", "C2")).toEqual(['=A6+Data!A5+LOG10(A6)&"A5"']);
  });

  test("other sheets that point here are rewritten, quoted names too", () => {
    const wb = book();
    wb.addSheet("My Data");
    wb.writeRange("My Data", "A1", [["=Sheet1!A3", "='Sheet1'!A3:A4", "=A3"]], human);
    wb.insertRows("Sheet1", 1, 2);
    expect(formulas(wb, "My Data", "A1:C1")).toEqual(["=Sheet1!A5", "=Sheet1!A5:A6", "=A3"]);
    wb.insertRows("My Data", 1, 1);
    expect(formulas(wb, "My Data", "A2:C2")).toEqual(["=Sheet1!A5", "=Sheet1!A5:A6", "=A4"]);
  });

  test("deleted cells become #REF!, ranges shrink, dead ranges become #REF!", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[1], [2], [3], [4]], human);
    wb.writeRange("Sheet1", "C1", [["=A2", "=A3", "=SUM(A1:A3)", "=SUM(A2:A3)", "=SUM(A3:A4)"]], human);
    wb.deleteRows("Sheet1", 2, 2);
    expect(wb.readRange("Sheet1", "A1:A2").flat().map((c) => c.value)).toEqual([1, 4]);
    expect(formulas(wb, "Sheet1", "C1:G1")).toEqual(["=#REF!", "=#REF!", "=SUM(A1:A1)", "=SUM(#REF!)", "=SUM(A2:A2)"]);
  });

  test("a deleted formula cell leaves no key behind", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [["=1"], ["=2"]], human);
    wb.deleteRows("Sheet1", 1, 2);
    expect(wb.cellCount("Sheet1")).toBe(0);
  });

  test("frozen rows and column widths follow the structure", () => {
    const wb = book();
    wb.setFrozenRows("Sheet1", 2);
    wb.setColumnWidth("Sheet1", "C", 120);
    wb.insertRows("Sheet1", 1, 1);
    expect(wb.sheetInfo("Sheet1").frozenRows).toBe(3);
    wb.insertRows("Sheet1", 4, 1);
    expect(wb.sheetInfo("Sheet1").frozenRows).toBe(3);
    wb.deleteRows("Sheet1", 1, 2);
    expect(wb.sheetInfo("Sheet1").frozenRows).toBe(1);
    wb.insertCols("Sheet1", 2, 1);
    expect(wb.sheetInfo("Sheet1").columnWidths).toEqual({ D: 120 });
    wb.deleteCols("Sheet1", 4, 1);
    expect(wb.sheetInfo("Sheet1").columnWidths).toEqual({});
  });

  test("bad positions and counts throw", () => {
    const wb = book();
    for (const [at, count] of [[0, 1], [1, 0], [1.5, 1], [2_000_000, 1], [1, 2_000_000], [1, 10_001]]) {
      expect(() => wb.insertRows("Sheet1", at, count)).toThrow(/invalid/);
    }
  });

  test("an insert that would push a cell off the grid is refused and changes nothing", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1048576", [[7]], human);
    wb.writeRange("Sheet1", "B1", [["=A1048576"]], human);
    expect(() => wb.insertRows("Sheet1", 1, 1)).toThrow(/grid/);
    expect(wb.readRange("Sheet1", "A1048576")[0][0].value).toBe(7);
    expect(formulas(wb, "Sheet1", "B1")).toEqual(["=A1048576"]);
    wb.setColumnWidth("Sheet1", "XFD", 50);
    expect(() => wb.insertCols("Sheet1", 1, 1)).toThrow(/grid/);
    expect(wb.sheetInfo("Sheet1").columnWidths).toEqual({ XFD: 50 });
  });

  test("a ref pushed past the last row is clamped or becomes #REF!", () => {
    const wb = book();
    wb.writeRange("Sheet1", "B1", [["=A1048576", "=SUM(A2:A1048576)"]], human);
    wb.insertRows("Sheet1", 1, 1);
    expect(formulas(wb, "Sheet1", "B2:C2")).toEqual(["=#REF!", "=SUM(A3:A1048576)"]);
  });
});

describe("insert and delete columns", () => {
  test("cells and refs shift right and left", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[1, 2, "=SUM(A1:B1)"]], human);
    wb.insertCols("Sheet1", 2, 1);
    expect(formulas(wb, "Sheet1", "A1:D1")).toEqual([1, null, 2, "=SUM(A1:C1)"]);
    wb.deleteCols("Sheet1", 1, 1);
    expect(formulas(wb, "Sheet1", "A1:C1")).toEqual([null, 2, "=SUM(A1:B1)"]);
  });
});

describe("rename and delete sheets", () => {
  test("rename rewrites refs; delete turns them into #REF!", () => {
    const wb = book();
    wb.addSheet("Data");
    wb.writeRange("Sheet1", "A1", [["=Data!B2", "=SUM(Data!A1:A3)", "=A1"]], human);
    wb.renameSheet("Data", "Raw data");
    expect(formulas(wb, "Sheet1", "A1:C1")).toEqual(["='Raw data'!B2", "=SUM('Raw data'!A1:A3)", "=A1"]);
    wb.deleteSheet("Raw data");
    expect(formulas(wb, "Sheet1", "A1:C1")).toEqual(["=#REF!", "=SUM(#REF!)", "=A1"]);
    expect(() => wb.deleteSheet("Sheet1")).toThrow(/last sheet/);
  });

  test("rename keeps own-sheet refs unqualified and rejects duplicates", () => {
    const wb = book();
    wb.addSheet("Two");
    wb.writeRange("Sheet1", "A1", [["=A2+Sheet1!A2"]], human);
    wb.renameSheet("Sheet1", "Main");
    expect(formulas(wb, "Main", "A1")).toEqual(["=A2+Main!A2"]);
    expect(() => wb.renameSheet("Two", "main")).toThrow(/exists/);
    wb.renameSheet("Two", "TWO");
    expect(wb.sheetNames()).toEqual(["Main", "TWO"]);
  });
});
