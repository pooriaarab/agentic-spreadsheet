import { describe, expect, test } from "vitest";
import * as Y from "yjs";
import { Workbook, type Principal } from "./workbook.js";

// Failure modes, listed before the code:
// - a formula string lands as a value, or a numeric-looking string becomes a number
// - clearing a cell leaves a key behind, or reading empty cells creates keys
// - a write to a missing sheet or past the grid edge passes silently
// - an agent edit loses its author kind, or an unknown kind is accepted
// - duplicate sheet names (any case) or names with banned characters pass
// - a bad frozen-row count or column width passes
// - concurrent writes to different cells lose one, or encoded state does not round-trip

const human: Principal = { id: "u1", kind: "human" };
const agent: Principal = { id: "a1", kind: "agent_of_human" };

function book(): Workbook {
  const wb = new Workbook();
  wb.addSheet("Sheet1");
  return wb;
}

describe("writes and reads", () => {
  test("formulas, numbers, and numeric-looking strings stay distinct", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[1, "007", "=A1+1", true, "text"]], human);
    const [a, b, c, d, e] = wb.readRange("Sheet1", "A1:E1")[0];
    expect([a.value, b.value, c.formula, c.value, d.value, e.value]).toEqual([1, "007", "=A1+1", null, true, "text"]);
  });

  test("clearing removes the key, and reading never creates keys", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[1, 2]], human);
    wb.writeRange("Sheet1", "A1", [[null, ""]], human);
    wb.readRange("Sheet1", "A1:Z50");
    expect(wb.cellCount("Sheet1")).toBe(0);
    expect(wb.readRange("Sheet1", "B2:C3")).toHaveLength(2);
  });

  test("a cleared cell keeps its format", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[5]], human);
    wb.setFormat("Sheet1", "A1", { bold: true }, human);
    wb.writeRange("Sheet1", "A1", [[null]], human);
    expect(wb.readRange("Sheet1", "A1")[0][0]).toMatchObject({ value: null, format: { bold: true } });
  });

  test("bad sheet and bad ranges throw", () => {
    const wb = book();
    expect(() => wb.writeRange("Nope", "A1", [[1]], human)).toThrow(/no sheet/);
    expect(() => wb.readRange("Sheet1", "A0")).toThrow(/cell/);
    expect(() => wb.writeRange("Sheet1", "XFD1", [[1, 2]], human)).toThrow(/grid/);
  });

  test("authors keep their kind; a later human edit relabels; bad kinds throw", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[1]], agent);
    expect(wb.readRange("Sheet1", "A1")[0][0].author).toMatchObject({ id: "a1", kind: "agent_of_human", machine: true });
    wb.writeRange("Sheet1", "A1", [[2]], human);
    expect(wb.readRange("Sheet1", "A1")[0][0].author).toMatchObject({ kind: "human", machine: false });
    wb.writeRange("Sheet1", "A2", [[3]], { id: "t", kind: "user_token" });
    expect(wb.readRange("Sheet1", "A2")[0][0].author?.machine).toBe(true);
    expect(() => wb.writeRange("Sheet1", "A3", [[4]], { id: "x", kind: "root" as never })).toThrow(/principal/);
  });
});

describe("sheets", () => {
  test("frozen rows and column widths validate", () => {
    const wb = book();
    wb.setFrozenRows("Sheet1", 2);
    wb.setColumnWidth("Sheet1", "c", 120);
    expect(wb.sheetInfo("Sheet1")).toEqual({ name: "Sheet1", frozenRows: 2, columnWidths: { C: 120 } });
    expect(() => wb.setFrozenRows("Sheet1", -1)).toThrow(/frozen/);
    expect(() => wb.setColumnWidth("Sheet1", "A1", 5)).toThrow(/column/);
    expect(() => wb.setColumnWidth("Sheet1", "A", 0)).toThrow(/width/);
  });

  test("order, duplicates, and bad names", () => {
    const wb = book();
    wb.addSheet("Two");
    wb.addSheet("Zero", 0);
    expect(wb.sheetNames()).toEqual(["Zero", "Sheet1", "Two"]);
    expect(() => wb.addSheet("sheet1")).toThrow(/exists/);
    expect(() => wb.addSheet("")).toThrow(/name/);
    expect(() => wb.addSheet("a!b")).toThrow(/name/);
  });
});

describe("sync", () => {
  test("concurrent writes merge and state round-trips", () => {
    const a = book();
    const b = Workbook.fromState(a.encode());
    a.writeRange("Sheet1", "A1", [[1]], human);
    b.writeRange("Sheet1", "B1", [[2]], agent);
    Y.applyUpdate(a.doc, b.encode());
    Y.applyUpdate(b.doc, a.encode());
    const copy = Workbook.fromState(a.encode());
    for (const w of [a, b, copy]) {
      expect(w.readRange("Sheet1", "A1:B1")[0].map((c) => c.value)).toEqual([1, 2]);
    }
    expect(copy.readRange("Sheet1", "B1")[0][0].author?.machine).toBe(true);
  });
});
