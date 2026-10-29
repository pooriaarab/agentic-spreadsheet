import { describe, expect, test } from "vitest";
import * as Y from "yjs";
import { Engine } from "./engine.js";
import { Workbook, type Principal } from "./workbook.js";
import { column, fork, grid, race, snapshot } from "../testing/replicas.js";

// Failure modes, listed before the code.
// - a move drops a cell, a format, or an author, or leaves a copy of the row where it was
// - a move to a place inside the block, or right after it, changes the order
// - a formula that points at a moved row keeps its old address, or a range that spans the move breaks
// - a row inserted after a moved row follows the move, or one inserted at the old place is lost
// - a move of a column changes a width, or leaves the width at the old column
// - a bad row, count, or target is accepted
// - two writers move one row at once and it shows twice, or in different places on each copy
// - an order that holds an entry with no id, a loop, or an anchor that does not exist stops a read,
//   hangs, or hides the rows that are fine
// - an insert far down a sparse sheet writes an entry for each row above it
// - an entry whose key repeats, or is a base id, makes the order loop forever or show a row twice
// - a range keeps a row that moved out of it, drops rows that did not move, or turns to #REF! when
//   the moved row is deleted
// - the frozen line follows a moved row and freezes the rows it passes
// - two inserts at once push a cell past the last row, and then reads and exports throw
// - a write to one cell rebuilds the order of the whole sheet
// - a change to the order that arrives from another copy does not tell a listener

const human: Principal = { id: "u1", kind: "human" };
const agent: Principal = { id: "a1", kind: "agent_of_human" };

function book(): Workbook {
  const wb = new Workbook();
  wb.addSheet("Sheet1");
  wb.writeRange("Sheet1", "A1", [[1], [2], [3], [4], [5]], human);
  return wb;
}

describe("move rows", () => {
  test("a block moves down and up, with its cells, formats, and authors", () => {
    const wb = book();
    wb.writeRange("Sheet1", "B2", [["x"]], agent);
    wb.setFormat("Sheet1", "A2", { bold: true }, agent);
    wb.moveRows("Sheet1", 2, 2, 5);
    expect(column(wb, "A", 5)).toEqual([1, 4, 2, 3, 5]);
    expect(grid(wb, "B3:B3")).toEqual([["x"]]);
    expect(wb.readRange("Sheet1", "A3")[0][0]).toMatchObject({ format: { bold: true }, author: { id: "u1" }, formatAuthor: { id: "a1" } });
    wb.moveRows("Sheet1", 4, 1, 2);
    expect(column(wb, "A", 5)).toEqual([1, 3, 4, 2, 5]);
    expect(wb.cellCount("Sheet1")).toBe(6);
  });

  test("a target inside the block or right after it changes nothing", () => {
    const wb = book();
    for (const to of [2, 3, 4]) wb.moveRows("Sheet1", 2, 2, to);
    expect(column(wb, "A", 5)).toEqual([1, 2, 3, 4, 5]);
  });

  test("a move to the top, and past the last used row", () => {
    const wb = book();
    wb.moveRows("Sheet1", 5, 1, 1);
    expect(column(wb, "A", 5)).toEqual([5, 1, 2, 3, 4]);
    wb.moveRows("Sheet1", 1, 1, 9);
    expect(column(wb, "A", 9)).toEqual([1, 2, 3, 4, null, null, null, 5, null]);
  });

  test("a formula follows a moved row, and a range spanning the move stays whole", () => {
    const wb = book();
    wb.writeRange("Sheet1", "C1", [["=A4", "=SUM(A1:A5)", "=SUM(A2:A3)"]], human);
    wb.moveRows("Sheet1", 4, 1, 1);
    expect(grid(wb, "C2:E2")).toEqual([["=A1", "=SUM(A2:A5)", "=SUM(A3:A4)"]]);
    expect(new Engine(wb).values("Sheet1", "C2:E2")).toEqual([[4, 11, 5]]);
  });

  test("a row inserted after a moved row stays beside the move, and the old place keeps its row", () => {
    const wb = book();
    wb.moveRows("Sheet1", 5, 1, 1);
    wb.insertRows("Sheet1", 2, 1);
    expect(column(wb, "A", 7)).toEqual([5, null, 1, 2, 3, 4, null]);
    wb.insertRows("Sheet1", 7, 1);
    expect(column(wb, "A", 8)).toEqual([5, null, 1, 2, 3, 4, null, null]);
  });

  test("a deleted row stays deleted after a move around it", () => {
    const wb = book();
    wb.deleteRows("Sheet1", 3, 1);
    wb.moveRows("Sheet1", 1, 1, 4);
    expect(column(wb, "A", 5)).toEqual([2, 4, 1, 5, null]);
  });

  test("bad rows, counts, and targets throw", () => {
    const wb = book();
    for (const [from, count, to] of [[0, 1, 1], [1, 0, 1], [1.5, 1, 1], [1, 1, 0], [1, 1, 2_000_000], [1, 10_001, 1]]) {
      expect(() => wb.moveRows("Sheet1", from, count, to)).toThrow(/invalid/);
    }
  });
});

describe("ranges and frozen rows when a row moves", () => {
  function twelve(): Workbook {
    const wb = new Workbook();
    wb.addSheet("Sheet1");
    wb.writeRange("Sheet1", "A1", Array.from({ length: 12 }, (_, i) => [i + 1]), human);
    wb.writeRange("Sheet1", "C20", [["=SUM(A1:A10)"]], human);
    return wb;
  }

  test("a row that moves out of a range leaves it, and the rest stay in", () => {
    const wb = twelve();
    wb.moveRows("Sheet1", 1, 1, 12);
    expect(grid(wb, "C20:C20")).toEqual([["=SUM(A1:A9)"]]);
    expect(new Engine(wb).value("Sheet1", "C20")).toBe(54);
    wb.deleteRows("Sheet1", 11, 1);
    expect(grid(wb, "C19:C19")).toEqual([["=SUM(A1:A9)"]]);
  });

  test("a row that moves into a range joins it", () => {
    const wb = twelve();
    wb.moveRows("Sheet1", 12, 1, 5);
    expect(grid(wb, "C20:C20")).toEqual([["=SUM(A1:A11)"]]);
    expect(new Engine(wb).value("Sheet1", "C20")).toBe(67);
  });

  test("a move of the last row of a range to the top shrinks the range from below", () => {
    const wb = twelve();
    wb.moveRows("Sheet1", 10, 1, 1);
    expect(grid(wb, "C20:C20")).toEqual([["=SUM(A2:A10)"]]);
  });

  test("the frozen line stays with the rows that are left in the frozen block", () => {
    const wb = book();
    wb.setFrozenRows("Sheet1", 2);
    wb.moveRows("Sheet1", 2, 1, 5);
    expect(wb.sheetInfo("Sheet1").frozenRows).toBe(1);
  });
});

describe("move columns", () => {
  test("cells, formulas, and widths follow", () => {
    const wb = new Workbook();
    wb.addSheet("Sheet1");
    wb.writeRange("Sheet1", "A1", [[1, 2, 3, "=B1"]], human);
    wb.setColumnWidth("Sheet1", "B", 90);
    wb.moveCols("Sheet1", 2, 1, 1);
    expect(grid(wb, "A1:D1")).toEqual([[2, 1, 3, "=A1"]]);
    expect(wb.sheetInfo("Sheet1").columnWidths).toEqual({ A: 90 });
  });
});

describe("two writers move one row at once", () => {
  test.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("it shows once, in the same place on both copies (client ids %j)", async (ids) => {
    const state = book().encode();
    const [a, b] = [fork(state, ids[0]), fork(state, ids[1])];
    await race(a, b, (w) => w.moveRows("Sheet1", 3, 1, 1), (w) => w.moveRows("Sheet1", 3, 1, 5));
    expect(snapshot(a)).toEqual(snapshot(b));
    expect(column(a, "A", 6).filter((v) => v === 3)).toHaveLength(1);
    expect([...column(a, "A", 5)].sort()).toEqual([1, 2, 3, 4, 5]);
  });
});

describe("a stored order that is not well formed", () => {
  test("rows that are fine still show, and nothing hangs", () => {
    const wb = book();
    const order = (wb.doc.getArray("sheets").get(0) as Y.Map<unknown>).get("rowOrder") as Y.Array<unknown>;
    order.push([
      null,
      "text",
      { k: "x" },
      { k: "loop.1", id: "loop.1", a: "loop.2" },
      { k: "loop.2", id: "loop.2", a: "loop.1" },
      { k: "lost.1", id: "lost.1", a: "nowhere.9" },
      { k: "Bad Key", id: "Bad Key", a: "3" },
    ]);
    wb.insertRows("Sheet1", 3, 1);
    expect(column(wb, "A", 6)).toEqual([1, 2, null, 3, 4, 5]);
  });

  test("an entry whose key repeats, or is a base id, is ignored", () => {
    const wb = book();
    const order = (wb.doc.getArray("sheets").get(0) as Y.Map<unknown>).get("rowOrder") as Y.Array<unknown>;
    order.push([
      { k: "a", id: "a", a: "0" },
      { k: "a", id: "b", a: "a" },
      { k: "5", id: "5", a: "5" },
      { k: "5", id: "z", a: "0" },
    ]);
    expect(column(wb, "A", 6)).toEqual([null, 1, 2, 3, 4, 5]);
  });

  test("a change to the order from another copy tells a listener, and a frozen line does not", () => {
    const wb = book();
    const seen: unknown[] = [];
    wb.onChange((change) => seen.push(change));
    const sheet = wb.doc.getArray("sheets").get(0) as Y.Map<unknown>;
    sheet.set("frozen", "3");
    expect(seen).toEqual([]);
    sheet.set("rowOrder", new Y.Array());
    expect(seen).toEqual([{ all: true }]);
  });

  test("a cell key that does not name a row and a column is not a cell", () => {
    const wb = book();
    const cells = (wb.doc.getArray("sheets").get(0) as Y.Map<unknown>).get("cells") as Y.Map<unknown>;
    cells.set("nonsense:c", 1);
    cells.set("1|:c", 1);
    cells.set("|1:c", 1);
    expect(wb.cellCount("Sheet1")).toBe(5);
    expect(wb.extent("Sheet1")).toBe("A1:A5");
  });
});

describe("two inserts at once at the edge of the grid", () => {
  test.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("no read throws, and the cell is back when the rows go (client ids %j)", async (ids) => {
    const wb = book();
    wb.writeRange("Sheet1", "A1048570", [["far"]], human);
    const state = wb.encode();
    const [a, b] = [fork(state, ids[0]), fork(state, ids[1])];
    await race(a, b, (w) => w.insertRows("Sheet1", 1, 5), (w) => w.insertRows("Sheet1", 1, 5));
    expect(snapshot(a)).toEqual(snapshot(b));
    expect(a.extent("Sheet1")).toBe("A1:A15");
    expect(a.cellCount("Sheet1")).toBe(5);
    expect(() => a.readRange("Sheet1", "A1048570:A1048576")).not.toThrow();
    a.deleteRows("Sheet1", 1, 10);
    expect(grid(a, "A1048570:A1048570")).toEqual([["far"]]);
    expect(a.cellCount("Sheet1")).toBe(6);
  });
});

describe("a big order", () => {
  test("a write to one cell does not rebuild it", { timeout: 120_000 }, () => {
    const wb = book();
    for (let i = 0; i < 15; i += 1) wb.insertRows("Sheet1", 1, 10_000);
    const started = Date.now();
    for (let i = 0; i < 20; i += 1) {
      wb.writeRange("Sheet1", "B150010", [[i]], human);
      expect(grid(wb, "B150010:B150010")).toEqual([[i]]);
    }
    expect(Date.now() - started).toBeLessThan(1_000);
  });
});

describe("a sparse sheet", () => {
  test("an insert deep down writes one entry per new row, not one per row above", () => {
    const wb = book();
    wb.writeRange("Sheet1", "A1000000", [["far"]], human);
    wb.insertRows("Sheet1", 500_000, 3);
    expect(grid(wb, "A1000003:A1000003")).toEqual([["far"]]);
    expect(((wb.doc.getArray("sheets").get(0) as Y.Map<unknown>).get("rowOrder") as Y.Array<unknown>).length).toBe(3);
    wb.deleteRows("Sheet1", 500_000, 3);
    expect(grid(wb, "A1000000:A1000000")).toEqual([["far"]]);
  });
});
