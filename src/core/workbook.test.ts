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
// - two writers touch the same NEW cell (value and format, or two format fields) and one loses
// - a grid of 150,000 rows throws RangeError, or a write takes a number past the cell text limit
// - NaN, Infinity, an object, or an array is stored as a value; a format patch takes unknown keys,
//   a bad color, or a bad alignment; a format write covers more than 100,000 cells
// - a format edit restamps the cell's value author, so machine output reads as human output
// - a column width of Infinity or "5" passes
// - clearing the content and the last format field leaves a key behind

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

describe("authors of value and format", () => {
  test("a format edit leaves the value author alone and records the format author", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[5]], agent);
    wb.setFormat("Sheet1", "A1", { numberFormat: "$0.00" }, human);
    expect(wb.readRange("Sheet1", "A1")[0][0]).toMatchObject({
      author: { id: "a1", machine: true },
      formatAuthor: { id: "u1", machine: false },
    });
  });

  test("a format-only cell has no value author; clearing a value keeps the format author", () => {
    const wb = book();
    wb.setFormat("Sheet1", "A1", { bold: true }, human);
    expect(wb.readRange("Sheet1", "A1")[0][0]).toMatchObject({ author: null, formatAuthor: { id: "u1" } });
    wb.writeRange("Sheet1", "A1", [[1]], agent);
    wb.writeRange("Sheet1", "A1", [[null]], agent);
    expect(wb.readRange("Sheet1", "A1")[0][0]).toMatchObject({ formatAuthor: { machine: false }, format: { bold: true } });
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
    for (const bad of [Infinity, "5", Number.NaN]) expect(() => wb.setColumnWidth("Sheet1", "A", bad as never)).toThrow(/width/);
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

  // Two copies start from the same state, edit offline, then swap updates.
  function fork(): [Workbook, Workbook] {
    const a = book();
    return [a, Workbook.fromState(a.encode())];
  }
  function merge(a: Workbook, b: Workbook): void {
    Y.applyUpdate(a.doc, Y.encodeStateAsUpdate(b.doc));
    Y.applyUpdate(b.doc, Y.encodeStateAsUpdate(a.doc));
  }

  test("a value and a format on the same new cell both survive", () => {
    const [a, b] = fork();
    a.writeRange("Sheet1", "A1", [[5]], human);
    b.setFormat("Sheet1", "A1", { bold: true }, agent);
    merge(a, b);
    for (const w of [a, b]) expect(w.readRange("Sheet1", "A1")[0][0]).toMatchObject({ value: 5, format: { bold: true } });
  });

  test("two format fields on the same new cell both survive", () => {
    const [a, b] = fork();
    a.setFormat("Sheet1", "A1", { bold: true }, human);
    b.setFormat("Sheet1", "A1", { color: "#ff0000" }, agent);
    merge(a, b);
    for (const w of [a, b]) expect(w.readRange("Sheet1", "A1")[0][0].format).toEqual({ bold: true, color: "#ff0000" });
  });

  test("a value and a formula written at once leave one winner, never both", () => {
    const [a, b] = fork();
    a.writeRange("Sheet1", "A1", [[5]], human);
    b.writeRange("Sheet1", "A1", [["=1+1"]], agent);
    merge(a, b);
    const [x, y] = [a, b].map((w) => w.readRange("Sheet1", "A1")[0][0]);
    expect(x).toEqual(y);
    expect([x.value, x.formula].filter((v) => v !== null)).toHaveLength(1);
  });
});

describe("limits and input checks", () => {
  test("a 150,000-row grid writes without a RangeError", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", Array.from({ length: 150_000 }, (_, i) => [i + 1]), human);
    expect(wb.cellCount("Sheet1")).toBe(150_000);
    expect(wb.readRange("Sheet1", "A150000")[0][0].value).toBe(150_000);
  });

  test.each([[Number.NaN], [Infinity], [{ a: 1 }], [[1]], ["x".repeat(32_768)], [`=${"1+".repeat(5_000)}1`]])(
    "a bad value %j throws",
    (value) => {
      expect(() => book().writeRange("Sheet1", "A1", [[value as never]], human)).toThrow(/value|text|formula/);
    },
  );

  test.each([
    [{ weight: true }, /format/], [{ color: "red; x" }, /color/], [{ background: "#12" }, /color/],
    [{ align: "middle" }, /align/], [{ bold: "yes" }, /bold/], [{ numberFormat: 5 }, /numberFormat/],
  ])("format patch %j throws", (patch, message) => {
    expect(() => book().setFormat("Sheet1", "A1", patch as never, human)).toThrow(message);
  });

  test("a format write over 100,000 cells throws", () => {
    expect(() => book().setFormat("Sheet1", "A1:Z4001", { bold: true }, human)).toThrow(/too large/);
  });

  test("clearing the content and the last format field leaves no key", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1", [[1]], human);
    wb.setFormat("Sheet1", "A1", { bold: true, align: "left" }, human);
    wb.setFormat("Sheet1", "A1", { bold: null, align: null }, human);
    wb.writeRange("Sheet1", "A1", [[null]], human);
    expect(wb.cellCount("Sheet1")).toBe(0);
  });
});
