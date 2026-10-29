import * as Y from "yjs";
import { Engine } from "../core/engine.js";
import { Workbook } from "../core/workbook.js";
import { interleaving } from "./interleaving.js";

// A copy of a workbook with its own client id. A fixed id makes a clash between two writers
// resolve the same way on every run.
export function fork(state: Uint8Array, clientID: number, gc = true): Workbook {
  const doc = new Y.Doc({ gc });
  doc.clientID = clientID;
  Y.applyUpdate(doc, state);
  return new Workbook(doc);
}

// Send what `to` lacks from `from`.
export function deliver(from: Workbook, to: Workbook): void {
  Y.applyUpdate(to.doc, Y.encodeStateAsUpdate(from.doc, Y.encodeStateVector(to.doc)));
}

export function exchange(a: Workbook, b: Workbook): void {
  deliver(a, b);
  deliver(b, a);
}

// Run `first` on `a` and `second` on `b` while neither copy has seen the other's change. A's change
// is in flight to B, and B acts before it lands. Then both copies exchange everything.
export async function race(a: Workbook, b: Workbook, first: (wb: Workbook) => void, second: (wb: Workbook) => void): Promise<void> {
  const lane = interleaving();
  const wire = lane.checkpoint("A's change travels to B");
  first(a);
  const inFlight = lane.start("A's change reaches B", () => wire.wrap(async () => deliver(a, b))());
  await lane.whileParkedAt(wire, () => {
    second(b);
    lane.mark("B edited before A's change landed");
  });
  await inFlight;
  exchange(a, b);
}

// What a reader sees: formula text, or the value, for each cell of a range.
export function grid(wb: Workbook, range: string, sheet = "Sheet1"): (string | number | boolean | null)[][] {
  return wb.readRange(sheet, range).map((row) => row.map((c) => c.formula ?? c.value));
}

export function column(wb: Workbook, letter: string, rows: number, sheet = "Sheet1"): (string | number | boolean | null)[] {
  return grid(wb, `${letter}1:${letter}${rows}`, sheet).map((row) => row[0]);
}

// Everything a reader can see, as one comparable value.
export function snapshot(wb: Workbook, sheet = "Sheet1", range = "A1:J20"): unknown {
  return {
    names: wb.sheetNames(),
    info: wb.sheetInfo(sheet),
    cells: wb.readRange(sheet, range),
    values: new Engine(wb).values(sheet, range),
  };
}
