import { describe, expect, test } from "vitest";
import * as Y from "yjs";
import { Engine } from "./engine.js";
import { Workbook } from "./workbook.js";
import { legacyState, stampOf } from "../testing/legacy.js";
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
// - a key already in the new form, in a sheet that still lacks an id, is dropped by the conversion
// - two copies that each convert an old workbook merge with no sign, though each made its own order
//   for the sheets and the merge keeps one, so rows and edits are lost and two sheets can share an id
// - a workbook that was converted once, saved, and then shared is flagged as a clash
// - a sheet that is converted and a sheet that is added later get one id, so a formula reads the
//   wrong sheet
// - old state that arrives after the workbook is made is read as empty and then overwritten
// - after a clash a read shows rows in the wrong place, as if the data were good
// - a conversion drops an old structure log that holds an unaccepted fork, so cells that may sit at
//   the wrong address read as good

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

  test.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("a workbook converted once and then shared keeps every edit (client ids %j)", async (ids) => {
    const converted = Workbook.fromState(legacyState()).encode();
    const [a, b] = [fork(converted, ids[0]), fork(converted, ids[1])];
    a.writeRange("Sheet1", "A1", [[555]], { id: "u2", kind: "human" });
    exchange(a, b);
    expect(column(b, "A", 1)).toEqual([555]);
    await race(a, b, (w) => (w.insertRows("Sheet1", 1, 1), w.writeRange("Sheet1", "A1", [["new"]], { id: "u2", kind: "human" })), (w) => w.writeRange("Sheet1", "A2", [[99]], { id: "u3", kind: "human" }));
    expect(snapshot(a)).toEqual(snapshot(b));
    expect(column(a, "A", 4)).toEqual(["new", 555, 99, "text"]);
    expect([a.conversionClash(), b.conversionClash()]).toEqual([false, false]);
  });

  test.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("copies that each convert one old state are flagged when they merge (client ids %j)", (ids) => {
    const state = legacyState();
    const [a, b] = [fork(state, ids[0]), fork(state, ids[1])];
    expect([a.conversionClash(), b.conversionClash()]).toEqual([false, false]);
    exchange(a, b);
    for (const w of [a, b]) {
      expect(w.conversionClash()).toBe(true);
      expect(() => w.sheetNames()).toThrow(/converted/);
    }
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

  test.each([[[1, 2]], [[2, 1]]] as [[number, number]][])("copies that convert different old states are flagged too, and every read and write throws (client ids %j)", (ids) => {
    const [older, newer] = twoStates();
    const [a, b] = [fork(older, ids[0]), fork(newer, ids[1])];
    expect([a.conversionClash(), b.conversionClash()]).toEqual([false, false]);
    exchange(a, b);
    for (const w of [a, b]) {
      expect(w.conversionClash()).toBe(true);
      expect(() => grid(w, "E9:E9")).toThrow(/converted/);
      expect(() => w.cellCount("Sheet1")).toThrow(/converted/);
      expect(() => w.sheetNames()).toThrow(/converted/);
      expect(() => w.writeRange("Sheet1", "A1", [[1]], { id: "u1", kind: "human" })).toThrow(/converted/);
      expect(() => w.insertRows("Sheet1", 1, 1)).toThrow(/converted/);
      const ids = (w.doc.getArray("sheets").toJSON() as { id: string; name: string }[]).map((sheet) => sheet.id);
      expect(new Set(ids).size).toBe(ids.length);
    }
  });

  test("a sheet added after 1,008 new ids does not share an id with a converted sheet", () => {
    const wb = Workbook.fromState(legacyState());
    wb.insertRows("Sheet1", 1, 1008);
    wb.addSheet("New");
    const ids = (wb.doc.getArray("sheets").toJSON() as { id: string }[]).map((sheet) => sheet.id);
    expect(new Set(ids).size).toBe(ids.length);
    wb.writeRange("New", "A1", [[42]], { id: "u1", kind: "human" });
    wb.writeRange("Sheet1", "H1", [["=New!A1"]], { id: "u1", kind: "human" });
    expect(grid(wb, "H1:H1")).toEqual([["=New!A1"]]);
    expect(new Engine(wb).value("Sheet1", "H1")).toBe(42);
  });

  test("old state that arrives after the workbook is made is refused, not read as empty", () => {
    const wb = new Workbook(new Y.Doc());
    Y.applyUpdate(wb.doc, legacyState());
    expect(() => wb.readRange("Sheet1", "A1:A1")).toThrow(/old A1 form/);
    expect(() => wb.writeRange("Sheet1", "A1", [["new"]], { id: "u1", kind: "human" })).toThrow(/old A1 form/);
    expect(() => wb.sheetNames()).toThrow(/old A1 form/);
    expect(grid(Workbook.fromState(wb.encode()), "A1:A1")).toEqual([[10]]);
  });

  test("a log that holds an unaccepted fork is refused, and an accepted one converts", () => {
    const fork2 = [{ id: "a", parent: null }, { id: "b", parent: null }];
    expect(() => Workbook.fromState(legacyState(fork2))).toThrow(/at the same time and not accepted/);
    expect(() => Workbook.fromState(legacyState([...fork2, { id: "c", parent: "b", accepted: true }]))).not.toThrow();
    expect(() => Workbook.fromState(legacyState([{ id: "a", parent: null }, { id: "b", parent: "a" }]))).not.toThrow();
    expect(() => Workbook.fromState(legacyState([]))).not.toThrow();
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

// Failure modes for the conversion guards, listed before the code:
// - a sheet with an id that holds an old-form key reads that cell as empty, with no error
// - an old-form key that arrives later, from a merge, is skipped the same way
// - a converted workbook that gets a sheet without an id converts to a clash when it loads again
// - an old formula with U+0001 or U+0002 reads as a ref after the conversion
describe("conversion guards", () => {
  const oldSheet = () => {
    const doc = new Y.Doc();
    Y.applyUpdate(doc, legacyState());
    return doc;
  };

  test("an old-form key on a sheet that has an id is an error", () => {
    const doc = new Y.Doc();
    Y.applyUpdate(doc, Workbook.fromState(legacyState()).encode());
    const cells = (doc.getArray("sheets").get(0) as Y.Map<unknown>).get("cells") as Y.Map<unknown>;
    cells.set("A1:c", "lost");
    const state = Y.encodeStateAsUpdate(doc);
    const wb = Workbook.fromState(state);
    expect(() => wb.readRange("Sheet1", "A1:A1")).toThrow(/cell key in the old A1 form/);
    expect(() => wb.sheetNames()).toThrow(/cell key in the old A1 form/);
  });

  test("an old-form key that arrives in a whole sheet or a replaced cells map is an error", () => {
    const sheetWith = (id: boolean) => {
      const doc = new Y.Doc();
      const sheet = new Y.Map<unknown>();
      if (id) sheet.set("id", "z:s9");
      sheet.set("name", "Far");
      sheet.set("cells", new Y.Map([["A1:c", 7]]));
      sheet.set("widths", new Y.Map());
      doc.getArray("sheets").push([sheet]);
      return Y.encodeStateAsUpdate(doc);
    };
    const whole = Workbook.fromState(legacyState());
    Y.applyUpdate(whole.doc, sheetWith(true));
    expect(() => whole.sheetNames()).toThrow(/cell key in the old A1 form/);
    const replaced = Workbook.fromState(legacyState());
    const other = new Y.Doc();
    Y.applyUpdate(other, replaced.encode());
    (other.getArray("sheets").get(0) as Y.Map<unknown>).set("cells", new Y.Map([["A1:c", 7]]));
    Y.applyUpdate(replaced.doc, Y.encodeStateAsUpdate(other));
    expect(() => replaced.sheetNames()).toThrow(/cell key in the old A1 form/);
  });

  test("an old-form key that arrives after the load is an error", () => {
    const wb = Workbook.fromState(legacyState());
    const other = new Y.Doc();
    Y.applyUpdate(other, wb.encode());
    ((other.getArray("sheets").get(0) as Y.Map<unknown>).get("cells") as Y.Map<unknown>).set("B9:c", "lost");
    Y.applyUpdate(wb.doc, Y.encodeStateAsUpdate(other));
    expect(() => wb.readRange("Sheet1", "A1:A1")).toThrow(/old A1 form/);
  });

  test("a converted workbook that gets a sheet without an id converts again and does not clash", () => {
    const wb = Workbook.fromState(legacyState());
    const extra = new Y.Doc();
    const sheet = new Y.Map<unknown>();
    sheet.set("name", "Late");
    const cells = new Y.Map<unknown>();
    cells.set("A1:c", 7);
    cells.set("A1:by", stampOf("u1", "human"));
    sheet.set("cells", cells);
    sheet.set("widths", new Y.Map());
    extra.getArray("sheets").push([sheet]);
    Y.applyUpdate(wb.doc, Y.encodeStateAsUpdate(extra));
    expect(() => wb.sheetNames()).toThrow(/load it with Workbook.fromState so it converts/);
    const again = Workbook.fromState(wb.encode());
    expect(again.conversionClash()).toBe(false);
    expect(grid(again, "A1:A1", "Late")).toEqual([[7]]);
    expect(grid(again, "A1:A1")).toEqual([[10]]);
  });

  test("two copies that each convert the same late sheet are a clash, and one copy is not", () => {
    const base = Workbook.fromState(legacyState());
    const extra = new Y.Doc();
    const sheet = new Y.Map<unknown>();
    sheet.set("name", "Late");
    sheet.set("cells", new Y.Map([["A1:c", 7], ["A1:by", stampOf("u1", "human")]]));
    sheet.set("widths", new Y.Map());
    extra.getArray("sheets").push([sheet]);
    Y.applyUpdate(base.doc, Y.encodeStateAsUpdate(extra));
    const saved = base.encode();
    const [one, two] = [1, 2].map((client) => {
      const doc = new Y.Doc();
      doc.clientID = 100 + client;
      Y.applyUpdate(doc, saved);
      return new Workbook(doc);
    });
    expect([one.conversionClash(), two.conversionClash()]).toEqual([false, false]);
    Y.applyUpdate(one.doc, Y.encodeStateAsUpdate(two.doc));
    expect(one.conversionClash()).toBe(true);
    expect(() => one.readRange("Late", "A1:A1")).toThrow(/two copies converted/);
  });

  test("a late sheet with no name is refused, and names with a colon are told apart", () => {
    const doc = new Y.Doc();
    Y.applyUpdate(doc, Workbook.fromState(legacyState()).encode());
    const sheet = new Y.Map<unknown>();
    sheet.set("cells", new Y.Map());
    doc.getArray("sheets").push([sheet]);
    expect(() => new Workbook(doc)).toThrow(/has no name/);
    const clash = new Workbook(new Y.Doc());
    clash.doc.getMap("conversions").set("late:a:x", true);
    clash.doc.getMap("conversions").set("late:a:b:y", true);
    expect(clash.conversionClash()).toBe(false);
  });

  test("an old formula with a control character is refused, not changed", () => {
    for (const char of ["\u0001", "\u0002"]) {
      const doc = oldSheet();
      const cells = (doc.getArray("sheets").get(0) as Y.Map<unknown>).get("cells") as Y.Map<unknown>;
      cells.set("E1:c", `="a${char}b"&A1`);
      expect(() => Workbook.fromState(Y.encodeStateAsUpdate(doc))).toThrow(/control character/);
    }
  });
});
