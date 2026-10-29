import { describe, expect, test } from "vitest";
import * as Y from "yjs";
import { Engine } from "./engine.js";
import { Workbook } from "./workbook.js";
import { column, exchange, fork, grid, race, snapshot } from "../testing/replicas.js";

// Failure modes, listed before the code. The fixture is built by hand, key for key, the way the
// version that keyed cells by A1 address wrote it.
// - a cell, a formula, an author, a format, a width, or the frozen rows changes on the way in
// - a formula that points at another sheet, a quoted sheet, or an absolute ref loses its target
// - a key in a form that cannot be read stays behind, or blocks the load
// - the old structure log stays, so a copy still sees a clash
// - two copies that both migrate disagree, or merging them adds or drops a cell
// - a migrated workbook still loses a write that races an insert
// - migrating a second time changes the workbook
// - two copies that each convert, and then edit before they merge, lose a value, an insert, or a write
// - two copies that convert different old states keep only their own cells after they merge
// - a key already in the new form, in a sheet that still lacks an id, is dropped by the conversion
// - two copies that convert different old states merge with no sign, then lose rows or edits, or
//   give two sheets one id so a formula reads the wrong sheet
// - two copies that convert one state are flagged as a clash because they keep deleted items
//   differently, and then no copy can be written again
// - after a clash a read shows rows in the wrong place, as if the data were good

const stampOf = (id: string, kind: string) => ({ id, kind, at: "2026-01-02T03:04:05.000Z" });

function legacyState(): Uint8Array {
  const doc = new Y.Doc();
  const sheets = doc.getArray<Y.Map<unknown>>("sheets");
  const make = (name: string, cells: Record<string, unknown>, widths: Record<string, number>, frozenRows: number) => {
    const sheet = new Y.Map<unknown>();
    sheet.set("name", name);
    const map = new Y.Map<unknown>();
    for (const [key, value] of Object.entries(cells)) map.set(key, value);
    sheet.set("cells", map);
    const w = new Y.Map<number>();
    for (const [letters, width] of Object.entries(widths)) w.set(letters, width);
    sheet.set("widths", w);
    sheet.set("frozenRows", frozenRows);
    sheets.push([sheet]);
  };
  const human = stampOf("u1", "human");
  const agent = stampOf("a1", "agent_of_human");
  make(
    "Sheet1",
    {
      "A1:c": 10, "A1:by": human,
      "A2:c": 20, "A2:by": agent,
      "A3:c": "text", "A3:by": human, "A3:fmt.bold": true, "A3:fby": agent,
      "B1:c": "=SUM(A1:A3)+'Q 3'!B2", "B1:by": human,
      "B2:c": "=$A$1*A2&\"A1\"", "B2:by": human,
      "C5:c": "=A2", "C5:by": agent,
      "D1:fmt.background": "#ff0000", "D1:fby": human,
      "bad:c": "oops", "ZZZZ9:c": 1,
    },
    { C: 120, B: 80 },
    2,
  );
  make("Q 3", { "B2:c": 5, "B2:by": human, "A1:c": "=Sheet1!A2", "A1:by": human }, {}, 0);
  doc.getArray("structure").push([{ id: "x", parent: null }, { id: "y", parent: "x" }]);
  return Y.encodeStateAsUpdate(doc);
}

describe("a workbook saved with A1 keys", () => {
  test("reads the same after the load", () => {
    const wb = Workbook.fromState(legacyState());
    expect(grid(wb, "A1:D3")).toEqual([[10, "=SUM(A1:A3)+'Q 3'!B2", null, null], [20, '=$A$1*A2&"A1"', null, null], ["text", null, null, null]]);
    expect(grid(wb, "C5:C5")).toEqual([["=A2"]]);
    expect(grid(wb, "A1:A1", "Q 3")).toEqual([["=Sheet1!A2"]]);
    const cells = wb.readRange("Sheet1", "A1:D3");
    expect(cells[1][0].author).toMatchObject({ id: "a1", kind: "agent_of_human", machine: true, at: "2026-01-02T03:04:05.000Z" });
    expect(cells[2][0]).toMatchObject({ format: { bold: true }, author: { id: "u1" }, formatAuthor: { id: "a1" } });
    expect(cells[0][3]).toMatchObject({ format: { background: "#ff0000" }, formatAuthor: { id: "u1" } });
    expect(wb.sheetInfo("Sheet1")).toEqual({ name: "Sheet1", frozenRows: 2, columnWidths: { B: 80, C: 120 } });
    expect(new Engine(wb).values("Sheet1", "A1:B2")).toEqual([[10, 35], [20, "200A1"]]);
    expect(new Engine(wb).value("Q 3", "A1")).toBe(20);
  });

  test("drops keys that are not addresses and the old structure log", () => {
    const wb = Workbook.fromState(legacyState());
    expect(wb.doc.getArray("structure").length).toBe(0);
    expect(wb.cellCount("Sheet1")).toBe(7);
    expect(wb.extent("Sheet1")).toBe("A1:D5");
  });

  test("a migrated workbook takes structure edits as a new one does", () => {
    const wb = Workbook.fromState(legacyState());
    wb.insertRows("Sheet1", 2, 1);
    expect(grid(wb, "A1:B3")).toEqual([[10, "=SUM(A1:A4)+'Q 3'!B2"], [null, null], [20, '=$A$1*A3&"A1"']]);
    expect(grid(wb, "A1:A1", "Q 3")).toEqual([["=Sheet1!A3"]]);
    wb.renameSheet("Q 3", "Quarter");
    expect(grid(wb, "B1:B1")).toEqual([["=SUM(A1:A4)+Quarter!B2"]]);
  });

  test("a second load of the saved copy changes nothing", () => {
    const once = Workbook.fromState(legacyState());
    const twice = Workbook.fromState(once.encode());
    expect(snapshot(twice)).toEqual(snapshot(once));
    expect(twice.doc.getArray("sheets").toJSON()).toEqual(once.doc.getArray("sheets").toJSON());
  });

  test("two copies that both migrate merge into the same workbook", () => {
    const state = legacyState();
    const a = fork(state, 1);
    const b = fork(state, 2);
    const before = a.doc.getArray("sheets").toJSON();
    exchange(a, b);
    expect(a.doc.getArray("sheets").toJSON()).toEqual(before);
    expect(b.doc.getArray("sheets").toJSON()).toEqual(before);
    expect(snapshot(a)).toEqual(snapshot(b));
  });

  test.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("copies that each convert and then edit keep every edit (client ids %j)", async (ids) => {
    const state = legacyState();
    const [a, b] = [fork(state, ids[0]), fork(state, ids[1])];
    a.writeRange("Sheet1", "A1", [[555]], { id: "u2", kind: "human" });
    exchange(a, b);
    expect(column(b, "A", 1)).toEqual([555]);
    await race(a, b, (w) => (w.insertRows("Sheet1", 1, 1), w.writeRange("Sheet1", "A1", [["new"]], { id: "u2", kind: "human" })), (w) => w.writeRange("Sheet1", "A2", [[99]], { id: "u3", kind: "human" }));
    expect(snapshot(a)).toEqual(snapshot(b));
    expect(column(a, "A", 4)).toEqual(["new", 555, 99, "text"]);
  });

  function twoStates(): [Uint8Array, Uint8Array] {
    const older = legacyState();
    const doc = new Y.Doc();
    Y.applyUpdate(doc, older);
    const sheets = doc.getArray<Y.Map<unknown>>("sheets");
    const extra = (sheets.get(0).get("cells") as Y.Map<unknown>);
    extra.set("E9:c", 7);
    extra.set("E9:by", stampOf("u1", "human"));
    const added = new Y.Map<unknown>();
    added.set("name", "Newest");
    added.set("cells", new Y.Map<unknown>());
    added.set("widths", new Y.Map<unknown>());
    added.set("frozenRows", 0);
    sheets.insert(0, [added]);
    return [older, Y.encodeStateAsUpdate(doc)];
  }

  test.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("copies that convert different old states are flagged, and every read and write throws (client ids %j)", (ids) => {
    const [older, newer] = twoStates();
    const [a, b] = [fork(older, ids[0]), fork(newer, ids[1])];
    expect([a.conversionClash(), b.conversionClash()]).toEqual([false, false]);
    exchange(a, b);
    for (const w of [a, b]) {
      expect(w.conversionClash()).toBe(true);
      expect(() => grid(w, "E9:E9")).toThrow(/different old states/);
      expect(() => w.cellCount("Sheet1")).toThrow(/different old states/);
      expect(() => w.sheetNames()).toThrow(/different old states/);
      expect(() => w.writeRange("Sheet1", "A1", [[1]], { id: "u1", kind: "human" })).toThrow(/different old states/);
      expect(() => w.insertRows("Sheet1", 1, 1)).toThrow(/different old states/);
      const ids = (w.doc.getArray("sheets").toJSON() as { id: string; name: string }[]).map((sheet) => sheet.id);
      expect(new Set(ids).size).toBe(ids.length);
    }
  });

  test("copies that keep deleted items differently are not flagged", () => {
    const doc = new Y.Doc({ gc: false });
    Y.applyUpdate(doc, legacyState());
    const cells = (doc.getArray("sheets").get(0) as Y.Map<unknown>).get("cells") as Y.Map<unknown>;
    cells.set("A1:c", 11);
    cells.set("A1:c", 12);
    const state = Y.encodeStateAsUpdate(doc);
    const [a, b] = [fork(state, 1, false), fork(state, 2, true)];
    exchange(a, b);
    expect([a.conversionClash(), b.conversionClash()]).toEqual([false, false]);
    a.writeRange("Sheet1", "A1", [[1]], { id: "u1", kind: "human" });
    exchange(a, b);
    expect(column(b, "A", 1)).toEqual([1]);
  });

  test.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("copies that convert one state are not flagged (client ids %j)", (ids) => {
    const state = legacyState();
    const [a, b] = [fork(state, ids[0]), fork(state, ids[1])];
    exchange(a, b);
    expect([a.conversionClash(), b.conversionClash()]).toEqual([false, false]);
    a.writeRange("Sheet1", "A1", [[1]], { id: "u1", kind: "human" });
  });

  test("a key already in the new form is kept", () => {
    const doc = new Y.Doc();
    Y.applyUpdate(doc, legacyState());
    const cells = (doc.getArray("sheets").get(0) as Y.Map<unknown>).get("cells") as Y.Map<unknown>;
    cells.set("3|4:c", "kept");
    cells.set("3|4:by", stampOf("u1", "human"));
    const wb = Workbook.fromState(Y.encodeStateAsUpdate(doc));
    expect(grid(wb, "D3:D3")).toEqual([["kept"]]);
  });

  test("a write that races an insert follows its row", async () => {
    const state = Workbook.fromState(legacyState()).encode();
    const a = fork(state, 1);
    const b = fork(state, 2);
    await race(a, b, (w) => w.insertRows("Sheet1", 2, 1), (w) => w.writeRange("Sheet1", "A2", [[99]], { id: "u2", kind: "human" }));
    expect(snapshot(a)).toEqual(snapshot(b));
    expect(column(a, "A", 4)).toEqual([10, null, 99, "text"]);
    expect(grid(a, "C6:C6")).toEqual([["=A3"]]);
  });
});
