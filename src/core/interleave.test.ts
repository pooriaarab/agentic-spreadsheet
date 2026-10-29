import { describe, expect, test } from "vitest";
import { colLetters } from "./a1.js";
import { Engine } from "./engine.js";
import { Workbook, type Principal } from "./workbook.js";
import { column, exchange, fork, grid, race, snapshot } from "../testing/replicas.js";

// Failure modes, listed before the code. Each one runs with both client-id orders, because a
// clash between two writers resolves by client id.
// - a value write that races a row or column insert lands at the old address, on the wrong row
// - two inserts at the same index merge into one insert, or put the rows in a different order on each copy
// - a write to a row that another writer deleted lands on the row that moved up into its place
// - two deletes of the same or of overlapping rows remove a row that neither one named
// - an insert next to a delete is lost, or the delete takes the new row
// - a formula written while another writer inserts or deletes keeps its old address
// - a formula that points at a moved row keeps pointing at the old address
// - the two copies differ after they exchange every change, whatever the order of the changes
// - a write that follows its row is lost when the row is moved by the other writer
// - a random mix of writes, inserts, deletes, and moves puts a value on a row it was not written to

const human: Principal = { id: "u1", kind: "human" };
const agent: Principal = { id: "a1", kind: "agent_of_human" };

function seed(): Uint8Array {
  const wb = new Workbook();
  wb.addSheet("Sheet1");
  wb.writeRange("Sheet1", "A1", [[1, "a"], [2, "b"], [3, "c"]], human);
  return wb.encode();
}

function copies(ids: [number, number]): [Workbook, Workbook] {
  const state = seed();
  return [fork(state, ids[0]), fork(state, ids[1])];
}

function same(a: Workbook, b: Workbook, sheet = "Sheet1"): void {
  expect(snapshot(a, sheet)).toEqual(snapshot(b, sheet));
}

describe.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("writes and structure edits made at once (client ids %j)", (ids) => {
  test("a value write follows its row when another writer inserts above it", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.insertRows("Sheet1", 2, 1), (w) => w.writeRange("Sheet1", "A2", [["B"]], agent));
    same(a, b);
    expect(column(a, "A", 4)).toEqual([1, null, "B", 3]);
    expect(column(a, "B", 4)).toEqual(["a", null, "b", "c"]);
    expect(a.readRange("Sheet1", "A3")[0][0].author?.machine).toBe(true);
  });

  test("a value write that comes first still follows its row when the insert comes second", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.writeRange("Sheet1", "A2", [["B"]], agent), (w) => w.insertRows("Sheet1", 2, 1));
    same(a, b);
    expect(column(a, "A", 4)).toEqual([1, null, "B", 3]);
  });

  test("a value write follows its column when another writer inserts a column", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.insertCols("Sheet1", 1, 2), (w) => w.writeRange("Sheet1", "B3", [["C"]], human));
    same(a, b);
    expect(grid(a, "A3:D3")).toEqual([[null, null, 3, "C"]]);
  });

  test("two inserts at the same index both apply, in the same order on both copies", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.insertRows("Sheet1", 2, 1), (w) => w.insertRows("Sheet1", 2, 1));
    same(a, b);
    expect(column(a, "A", 5)).toEqual([1, null, null, 2, 3]);
    expect(column(a, "B", 5)).toEqual(["a", null, null, "b", "c"]);
  });

  test("two inserts of different sizes at the same index add up", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.insertRows("Sheet1", 2, 2), (w) => w.insertRows("Sheet1", 2, 3));
    same(a, b);
    expect(column(a, "A", 7)).toEqual([1, null, null, null, null, null, 2]);
    expect(column(a, "A", 8)[7]).toBe(3);
  });

  test("two inserts at different indexes both apply", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.insertRows("Sheet1", 2, 1), (w) => w.insertRows("Sheet1", 3, 1));
    same(a, b);
    expect(column(a, "A", 5)).toEqual([1, null, 2, null, 3]);
  });

  test("a write to a row that another writer deletes goes with the row", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.deleteRows("Sheet1", 2, 1), (w) => w.writeRange("Sheet1", "A2", [["B"]], agent));
    same(a, b);
    expect(column(a, "A", 3)).toEqual([1, 3, null]);
    expect(column(a, "B", 3)).toEqual(["a", "c", null]);
    expect(a.cellCount("Sheet1")).toBe(4);
  });

  test("a write to the row below a deleted row follows it up", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.deleteRows("Sheet1", 2, 1), (w) => w.writeRange("Sheet1", "A3", [["C"]], human));
    same(a, b);
    expect(column(a, "A", 3)).toEqual([1, "C", null]);
  });

  test("a write to a column that another writer deletes goes with the column", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.deleteCols("Sheet1", 1, 1), (w) => w.writeRange("Sheet1", "A2", [["B"]], agent));
    same(a, b);
    expect(grid(a, "A1:B3")).toEqual([["a", null], ["b", null], ["c", null]]);
  });

  test("deletes of the same row, and of overlapping rows, remove only the rows named", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.deleteRows("Sheet1", 2, 1), (w) => w.deleteRows("Sheet1", 2, 1));
    same(a, b);
    expect(column(a, "A", 3)).toEqual([1, 3, null]);
    const [c, d] = copies(ids);
    await race(c, d, (w) => w.deleteRows("Sheet1", 1, 2), (w) => w.deleteRows("Sheet1", 2, 2));
    same(c, d);
    expect(column(c, "A", 2)).toEqual([null, null]);
  });

  test("an insert next to a delete keeps the new row", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.deleteRows("Sheet1", 2, 1), (w) => w.insertRows("Sheet1", 2, 1));
    same(a, b);
    expect(column(a, "A", 4)).toEqual([1, null, 3, null]);
  });

  test("a formula written while another writer inserts keeps pointing at its row", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.insertRows("Sheet1", 2, 1), (w) => w.writeRange("Sheet1", "C1", [["=A3", "=SUM(A1:A3)"]], human));
    same(a, b);
    expect(grid(a, "C1:D1")).toEqual([["=A4", "=SUM(A1:A4)"]]);
    expect(new Engine(a).values("Sheet1", "C1:D1")).toEqual([[3, 6]]);
  });

  test("a formula written while another writer deletes its row becomes #REF!", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.deleteRows("Sheet1", 3, 1), (w) => w.writeRange("Sheet1", "C1", [["=A3", "=SUM(A1:A3)"]], human));
    same(a, b);
    expect(grid(a, "C1:D1")).toEqual([["=#REF!", "=SUM(A1:A2)"]]);
  });

  test("a formula follows a row that another writer moves", async () => {
    const [a, b] = copies(ids);
    b.writeRange("Sheet1", "C1", [["=A3"]], human);
    exchange(a, b);
    await race(a, b, (w) => w.moveRows("Sheet1", 3, 1, 1), (w) => w.insertRows("Sheet1", 2, 1));
    same(a, b);
    expect(column(a, "A", 4)).toEqual([3, 1, null, 2]);
    expect(grid(a, "C2:C2")).toEqual([["=A1"]]);
    expect(new Engine(a).value("Sheet1", "C2")).toBe(3);
  });

  test("a write follows a row that another writer moves", async () => {
    const [a, b] = copies(ids);
    await race(a, b, (w) => w.moveRows("Sheet1", 1, 1, 4), (w) => w.writeRange("Sheet1", "B1", [["A"]], human));
    same(a, b);
    expect(grid(a, "A1:B3")).toEqual([[2, "b"], [3, "c"], [1, "A"]]);
  });

  test("a frozen row and a column width follow the structure edits of the other writer", async () => {
    const [a, b] = copies(ids);
    a.setFrozenRows("Sheet1", 2);
    a.setColumnWidth("Sheet1", "B", 80);
    exchange(a, b);
    await race(a, b, (w) => w.insertRows("Sheet1", 1, 1), (w) => w.insertCols("Sheet1", 1, 1));
    same(a, b);
    expect(a.sheetInfo("Sheet1")).toMatchObject({ frozenRows: 3, columnWidths: { C: 80 } });
  });
});

describe("a merge never flags or refuses", () => {
  test("writes go on after two inserts at once, with no step to accept", async () => {
    const [a, b] = copies([1, 2]);
    await race(a, b, (w) => w.insertRows("Sheet1", 2, 1), (w) => w.insertRows("Sheet1", 2, 1));
    for (const w of [a, b]) w.writeRange("Sheet1", "A1", [[9]], human);
    exchange(a, b);
    same(a, b);
    expect(column(a, "A", 1)).toEqual([9]);
  });

  test("an insert and a rename made at once both apply", async () => {
    const [a, b] = copies([1, 2]);
    b.addSheet("Data");
    b.writeRange("Data", "A1", [["=Sheet1!A3"]], human);
    exchange(a, b);
    await race(a, b, (w) => w.insertRows("Sheet1", 2, 1), (w) => w.renameSheet("Sheet1", "Main"));
    same(a, b, "Main");
    expect(a.sheetNames()).toEqual(["Main", "Data"]);
    expect(grid(a, "A1:A1", "Data")).toEqual([["=Main!A4"]]);
  });
});

// A tiny seeded generator, so a failing run can be replayed.
function random(seedNumber: number): () => number {
  let s = seedNumber >>> 0;
  return () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = s;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// Column A holds a label on each of 8 rows. Each token goes in a cell of its own, in columns B and
// beyond. Whatever the writers insert, delete, or move, a token must stay on the row whose label it
// was written beside, and it must vanish with that row.
describe("a token stays on the row it was written to", () => {
  test.each(Array.from({ length: 40 }, (_, i) => i + 1))("seed %i", (seedNumber) => {
    const next = random(seedNumber);
    const pick = (n: number) => Math.floor(next() * n);
    const base = new Workbook();
    base.addSheet("Sheet1");
    base.writeRange("Sheet1", "A1", Array.from({ length: 8 }, (_, i) => [`row${i + 1}`]), human);
    const state = base.encode();
    const writers = [fork(state, 11), fork(state, 22), fork(state, 33)];
    const written = new Map<string, string>();
    let tokens = 0;

    const labelRow = (wb: Workbook, label: string): number => column(wb, "A", 40).indexOf(label) + 1;
    for (let step = 0; step < 14; step += 1) {
      const wb = writers[pick(writers.length)];
      const action = pick(6);
      const rows = 1 + pick(10);
      if (action <= 1) {
        const label = `row${1 + pick(8)}`;
        const at = labelRow(wb, label);
        if (at > 0) {
          const token = `t${tokens++}`;
          wb.writeRange("Sheet1", `${colLetters(2 + tokens - 1)}${at}`, [[token]], agent);
          written.set(token, label);
        }
      } else if (action === 2) wb.insertRows("Sheet1", rows, 1 + pick(2));
      else if (action === 3) wb.deleteRows("Sheet1", rows, 1);
      else if (action === 4) {
        const to = 1 + pick(12);
        if (to < rows || to > rows + 1) wb.moveRows("Sheet1", rows, 1, to);
      } else exchange(wb, writers[pick(writers.length)]);
    }
    exchange(writers[0], writers[1]);
    exchange(writers[1], writers[2]);
    exchange(writers[0], writers[1]);
    for (const wb of writers) expect(snapshot(wb)).toEqual(snapshot(writers[0]));

    const [wb] = writers;
    const labels = column(wb, "A", 40);
    const cells = grid(wb, "A1:Q40");
    for (const [token, label] of written) {
      const labelAt = labels.indexOf(label);
      const tokenAt = cells.findIndex((row) => row.includes(token));
      if (labelAt >= 0) expect(tokenAt, `${token} beside ${label}`).toBe(labelAt);
      else expect(tokenAt, `${token} outlives ${label}`).toBe(-1);
    }
  });
});
