import * as Y from "yjs";

export const stampOf = (id: string, kind: string) => ({ id, kind, at: "2026-01-02T03:04:05.000Z" });

// Build the state of a workbook the way the version that keyed cells by A1 address saved it, key
// for key. `log` is its structure log.
export function legacyState(log: unknown[] = [{ id: "x", parent: null }, { id: "y", parent: "x" }]): Uint8Array {
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
  doc.getArray("structure").push(log);
  return Y.encodeStateAsUpdate(doc);
}
