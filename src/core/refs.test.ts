import { describe, expect, test } from "vitest";
import { formatSpan, mapRefs, quoteSheet, readRef, sameSheet } from "./refs.js";

// Failure modes, listed before the tests:
// - $ flags are lost, so absolute refs turn relative when rewritten
// - a range is not normalised, or a lone cell is flagged as a range
// - quoted sheet names with spaces or '' lose the quote on the way back
// - text inside a string literal, or LOG10( , is rewritten as a ref
// - a ref glued to a word (XA1, B.A1, Sheet!A1 split in the middle) is read mid-word
// - quoteSheet leaves a name that looks like a cell unquoted

describe("readRef", () => {
  test("reads a cell and reports its text and flags", () => {
    const found = readRef("=A1+B2", 1, "Home");
    expect(found).toEqual({
      text: "A1",
      ref: { sheet: null, target: "Home", range: { c1: 1, r1: 1, c2: 1, r2: 1 } },
      flags: { c1: "", r1: "", c2: "", r2: "", isRange: false },
    });
  });

  test("keeps $ flags per part", () => {
    const found = readRef("$A1:B$3", 0, "H")!;
    expect(found.flags).toEqual({ c1: "$", r1: "", c2: "", r2: "$", isRange: true });
    expect(found.ref.range).toEqual({ c1: 1, r1: 1, c2: 2, r2: 3 });
  });

  test("returns null where no ref starts", () => {
    expect(readRef("hello", 0, "H")).toBeNull();
    expect(readRef("A1", 1, "H")).toBeNull();
    expect(readRef("LOG10(1)", 0, "H")).toBeNull();
    expect(readRef("A1(", 0, "H")).toBeNull();
    expect(readRef("A1x", 0, "H")).toBeNull();
  });

  test("returns null for cells outside the grid", () => {
    expect(readRef("XFE1", 0, "H")).toBeNull();
    expect(readRef("A0", 0, "H")).toBeNull();
    expect(readRef("A1048577", 0, "H")).toBeNull();
    expect(readRef("A1:A1048577", 0, "H")).toBeNull();
    expect(readRef("A1:A0", 0, "H")).toBeNull();
    expect(readRef("XFD1048576", 0, "H")).not.toBeNull();
  });

  test("normalises a reversed range but keeps the flags where they were written", () => {
    const found = readRef("$C$3:A1", 0, "H")!;
    expect(found.ref.range).toEqual({ c1: 1, r1: 1, c2: 3, r2: 3 });
    expect(found.flags).toMatchObject({ c1: "$", r1: "$", c2: "", r2: "" });
  });

  test("reads plain and quoted sheet names", () => {
    expect(readRef("Data!A1", 0, "H")!.ref).toMatchObject({ sheet: "Data", target: "Data" });
    expect(readRef("'My Sheet'!A1:B2", 0, "H")!.ref).toMatchObject({ sheet: "My Sheet", target: "My Sheet" });
    expect(readRef("'It''s'!A1", 0, "H")!.ref.sheet).toBe("It's");
  });

  test("a sheet-qualified range takes the sheet from the first part only", () => {
    const found = readRef("Data!A1:B2", 0, "H")!;
    expect(found.text).toBe("Data!A1:B2");
    expect(found.ref.range).toEqual({ c1: 1, r1: 1, c2: 2, r2: 2 });
  });
});

describe("formatSpan", () => {
  test("writes a lone cell with its flags", () => {
    expect(formatSpan({ c1: 28, r1: 5, c2: 28, r2: 5 }, { c1: "$", r1: "", c2: "", r2: "", isRange: false })).toBe("$AB5");
  });

  test("writes a range with a flag on each part", () => {
    expect(formatSpan({ c1: 1, r1: 1, c2: 3, r2: 9 }, { c1: "$", r1: "", c2: "", r2: "$", isRange: true })).toBe("$A1:C$9");
  });

  test("round-trips what readRef reads", () => {
    for (const text of ["A1", "$A$1", "A1:B2", "$A1:B$2", "$AA$10:$AB$20"]) {
      const { ref, flags } = readRef(text, 0, "H")!;
      expect(formatSpan(ref.range, flags)).toBe(text);
    }
  });
});

describe("mapRefs", () => {
  const upper = (found: { text: string }) => `<${found.text}>`;

  test("rewrites every ref and leaves the rest alone", () => {
    expect(mapRefs("=SUM(A1:B2)+$C$3*2", "H", upper)).toBe("=SUM(<A1:B2>)+<$C$3>*2");
  });

  test("leaves string literals alone, including doubled quotes", () => {
    expect(mapRefs('="A1"&A1&"x""B2"&B2', "H", upper)).toBe('="A1"&<A1>&"x""B2"&<B2>');
  });

  test("an unterminated literal runs to the end without a crash", () => {
    expect(mapRefs('=A1&"B2', "H", upper)).toBe('=<A1>&"B2');
  });

  test("does not read a ref in the middle of a word", () => {
    expect(mapRefs("=B.A1+_A1+1A1+$$A1", "H", upper)).toBe("=B.A1+_A1+1A1+$$A1");
    expect(mapRefs("=LOG10(A1)", "H", upper)).toBe("=LOG10(<A1>)");
  });

  test("reads sheet-qualified refs whole", () => {
    expect(mapRefs("='My Sheet'!A1+Data!B2:C3", "H", upper)).toBe("=<'My Sheet'!A1>+<Data!B2:C3>");
  });

  test("passes the home sheet and the parsed ref to edit", () => {
    const seen: string[] = [];
    mapRefs("=A1+Other!B2", "Home", ({ ref }) => {
      seen.push(ref.target);
      return "x";
    });
    expect(seen).toEqual(["Home", "Other"]);
  });

  test("a formula with no refs comes back the same", () => {
    expect(mapRefs("=1+2", "H", () => "never")).toBe("=1+2");
    expect(mapRefs("", "H", () => "never")).toBe("");
  });
});

describe("quoteSheet and sameSheet", () => {
  test("plain names stay plain", () => {
    expect(quoteSheet("Sheet1")).toBe("Sheet1");
    expect(quoteSheet("my.sheet_2")).toBe("my.sheet_2");
  });

  test("names with spaces, quotes, or a leading digit are quoted", () => {
    expect(quoteSheet("My Sheet")).toBe("'My Sheet'");
    expect(quoteSheet("It's")).toBe("'It''s'");
    expect(quoteSheet("2024")).toBe("'2024'");
    expect(quoteSheet("")).toBe("''");
  });

  test("a name that looks like a cell is quoted so it reads back as a sheet", () => {
    expect(quoteSheet("A1")).toBe("'A1'");
    expect(quoteSheet("abc123")).toBe("'abc123'");
    expect(readRef(`${quoteSheet("A1")}!B2`, 0, "H")!.ref.sheet).toBe("A1");
  });

  test("every quoted name reads back to itself", () => {
    for (const name of ["Sheet1", "My Sheet", "It's", "a!b", "A1", "x.y"]) {
      expect(readRef(`${quoteSheet(name)}!A1`, 0, "H")!.ref.sheet).toBe(name);
    }
  });

  test("sheet names compare without case", () => {
    expect(sameSheet("Data", "DATA")).toBe(true);
    expect(sameSheet("Data", "Date")).toBe(false);
  });
});
