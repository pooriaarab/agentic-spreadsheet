import { describe, expect, test } from "vitest";
import { colIndex, colLetters, formatCell, parseCell, parseRange } from "./a1.js";

// Failure modes, listed before the code:
// - letters and indexes drift at the Z/AA and XFD edges
// - "A0", "1A", "A", "", and columns past XFD parse as cells
// - a reversed range keeps its order and breaks loops
// - a lone cell does not parse as a range

describe("a1", () => {
  test("column letters round-trip at the edges", () => {
    for (const [letters, index] of [["A", 1], ["Z", 26], ["AA", 27], ["AZ", 52], ["BA", 53], ["XFD", 16384]] as const) {
      expect(colIndex(letters)).toBe(index);
      expect(colLetters(index)).toBe(letters);
    }
  });

  test("cells parse with and without $ and in lower case", () => {
    expect(parseCell("$b$2")).toEqual({ col: 2, row: 2 });
    expect(parseCell("AA10")).toEqual({ col: 27, row: 10 });
    expect(formatCell({ col: 27, row: 10 })).toBe("AA10");
  });

  test.each(["A0", "1A", "A", "", "XFE1", "A1048577", "A1:B2", " A1"])("%j is not a cell", (text) => {
    expect(() => parseCell(text)).toThrow(/cell/);
  });

  test("ranges normalise and accept a lone cell", () => {
    expect(parseRange("C3:A1")).toEqual({ c1: 1, r1: 1, c2: 3, r2: 3 });
    expect(parseRange("B2")).toEqual({ c1: 2, r1: 2, c2: 2, r2: 2 });
    expect(() => parseRange("A1:")).toThrow(/range/);
  });
});
