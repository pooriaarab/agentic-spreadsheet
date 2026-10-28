import { describe, expect, test } from "vitest";
import * as Y from "yjs";
import { Engine, type Value } from "./engine.js";
import { Workbook, type CellInput } from "./workbook.js";

// Failure modes, listed before the code:
// - a chain of 100,000 cells overflows the stack, returns an error, or takes quadratic time
// - a loop of 100,000 cells crashes, or one cell of the loop is not #CIRCULAR
// - SUM over 100,000 formula cells crashes or walks the range once per cell
// - a second read recomputes cells that did not change
// - an edit recomputes cells that do not depend on it, or skips one that does (direct, through a
//   range, through a blank cell that gets filled, through another sheet, after a row insert, after
//   an update that arrives from another copy of the workbook)
// - a format change recomputes anything
// - a loop that an edit breaks keeps its old #CIRCULAR, or a loop reads differently by start cell
// - IF(B1,1,A1) in A1 reports a loop while B1 is TRUE, or keeps the value after B1 turns FALSE
// - NOW() and a cell that reads it go stale between reads, or recompute cells that do not read it

const me = { id: "u", kind: "human" } as const;

function setup(cells: Record<string, CellInput> = {}) {
  const wb = new Workbook();
  wb.addSheet("Sheet1");
  for (const [at, v] of Object.entries(cells)) wb.writeRange("Sheet1", at, [[v]], me);
  return { wb, engine: new Engine(wb) };
}
const column = (wb: Workbook, at: string, n: number, formula: (i: number) => string) =>
  wb.writeRange("Sheet1", at, Array.from({ length: n }, (_, i) => [formula(i + 1)]), me);
const circular: Value = { error: "#CIRCULAR" };

describe("long chains", () => {
  test("a chain of 100,000 cells reads from its far end, once", { timeout: 120_000 }, () => {
    const { wb, engine } = setup({ A1: 1 });
    column(wb, "B1", 100_000, (i) => (i === 1 ? "=A1" : `=B${i - 1}+1`));
    expect(engine.value("Sheet1", "B100000")).toBe(100_000);
    expect(engine.evaluations).toBe(100_000);
    expect(engine.value("Sheet1", "B99999")).toBe(99_999);
    expect(engine.evaluations).toBe(100_000);
    wb.writeRange("Sheet1", "A1", [[2]], me);
    expect(engine.value("Sheet1", "B100000")).toBe(100_001);
    expect(engine.evaluations).toBe(200_000);
  });

  test("a loop of 100,000 cells is #CIRCULAR in every cell and does not crash", { timeout: 120_000 }, () => {
    const { wb, engine } = setup();
    column(wb, "A1", 100_000, (i) => (i === 1 ? "=A100000+1" : `=A${i - 1}+1`));
    expect(engine.value("Sheet1", "A50000")).toEqual(circular);
    expect(engine.values("Sheet1", "A1:A3")).toEqual([[circular], [circular], [circular]]);
    wb.writeRange("Sheet1", "A1", [[7]], me);
    expect(engine.value("Sheet1", "A100000")).toBe(100_006);
  });

  test("SUM over 100,000 formula cells", { timeout: 120_000 }, () => {
    const { wb, engine } = setup();
    column(wb, "A1", 100_000, () => "=1+0");
    wb.writeRange("Sheet1", "B1", [["=SUM(A1:A100000)"]], me);
    expect(engine.value("Sheet1", "B1")).toBe(100_000);
    expect(engine.evaluations).toBe(100_001);
    wb.writeRange("Sheet1", "A777", [["=5"]], me);
    expect(engine.value("Sheet1", "B1")).toBe(100_004);
    expect(engine.evaluations).toBe(100_003);
  });
});

describe("only what changed is recomputed", () => {
  const diamond = () => {
    const s = setup({ A1: 1, B1: "=A1+1", C1: "=A1+2", D1: "=B1+C1", E1: 10, F1: "=E1+1", G1: "=SUM(A1:A3)", H1: "=Other!A1" });
    s.wb.addSheet("Other");
    s.wb.writeRange("Other", "A1", [[4]], me);
    return s;
  };
  const all = (e: Engine) => e.values("Sheet1", "A1:H1");

  test("a second read recomputes nothing", () => {
    const { engine } = diamond();
    all(engine);
    const first = engine.evaluations;
    all(engine);
    expect([first, engine.evaluations]).toEqual([6, 6]);
  });

  test.each([
    ["a cell nothing reads", "Z9", 5, 0],
    ["the head of the diamond", "A1", 5, 4],
    ["the other branch", "E1", 20, 1],
    ["a blank cell inside a range", "A2", 3, 1],
    ["a blank cell that a range will cover", "A3", 3, 1],
    ["a cell on another sheet", "Other!A1", 9, 1],
  ])("an edit to %s recomputes exactly its readers", (_name, target, value, count) => {
    const { wb, engine } = diamond();
    all(engine);
    const before = engine.evaluations;
    const [sheet, cell] = target.includes("!") ? target.split("!") : ["Sheet1", target];
    wb.writeRange(sheet, cell, [[value]], me);
    all(engine);
    expect(engine.evaluations - before).toBe(count);
    const fresh = all(new Engine(wb));
    expect(all(engine)).toEqual(fresh);
  });

  test("a format change recomputes nothing", () => {
    const { wb, engine } = diamond();
    all(engine);
    const before = engine.evaluations;
    wb.setFormat("Sheet1", "A1:D1", { bold: true }, me);
    all(engine);
    expect(engine.evaluations).toBe(before);
  });

  test("a row insert, a rename, and an update from another copy give fresh values", () => {
    const { wb, engine } = diamond();
    all(engine);
    wb.insertRows("Sheet1", 1, 1);
    
    expect(engine.value("Sheet1", "D2")).toBe(5);
    wb.renameSheet("Other", "Elsewhere");
    expect(engine.value("Sheet1", "H2")).toBe(4);
    const copy = Workbook.fromState(wb.encode());
    copy.writeRange("Sheet1", "A2", [[100]], me);
    Y.applyUpdate(wb.doc, Y.encodeStateAsUpdate(copy.doc));
    expect(engine.value("Sheet1", "D2")).toBe(203);
  });
});

describe("loops and branches", () => {
  test("a loop that IFERROR wraps reads the same from either cell, and an edit breaks it", () => {
    const cells = { A1: "=IFERROR(B1,5)", B1: "=A1" };
    const [x, y] = [setup(cells), setup(cells)];
    expect([x.engine.value("Sheet1", "B1"), y.engine.values("Sheet1", "A1:B1")]).toEqual([circular, [[circular, circular]]]);
    x.wb.writeRange("Sheet1", "B1", [[3]], me);
    expect(x.engine.values("Sheet1", "A1:B1")).toEqual([[3, 3]]);
  });

  test("a ref in a branch that does not run is not a dependency until the branch runs", () => {
    const { wb, engine } = setup({ A1: "=IF(B1,1,A1)", B1: true });
    expect(engine.value("Sheet1", "A1")).toBe(1);
    wb.writeRange("Sheet1", "B1", [[false]], me);
    expect(engine.value("Sheet1", "A1")).toEqual(circular);
    wb.writeRange("Sheet1", "B1", [[true]], me);
    expect(engine.value("Sheet1", "A1")).toBe(1);
  });

  test("NOW and its readers refresh on every read, and nothing else does", async () => {
    const { engine } = setup({ A1: "=NOW()", B1: "=A1+0", C1: 5, D1: "=C1+1" });
    const read = () => engine.values("Sheet1", "A1:D1")[0];
    const [first, base] = [read(), engine.evaluations];
    await new Promise((resolve) => setTimeout(resolve, 20));
    const second = read();
    expect(second[0] as number).toBeGreaterThan(first[0] as number);
    expect(second[1] as number).toBeGreaterThan(first[1] as number);
    expect([second[3], engine.evaluations - base]).toEqual([6, 2]);
  });
});
