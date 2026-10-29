import { describe, expect, test } from "vitest";
import { FormatError, NUMBER_FORMATS } from "agentic-formats";
import type { Workbook as FileBook } from "agentic-formats";
import { Engine } from "./core/engine.js";
import { Workbook, type Principal } from "./core/workbook.js";
import { addSheets, exportFile, formatOf, importFile, MAX_IMPORT_SHEETS, MIME, toFileBook, type SheetFormat } from "./formats.js";

// Failure modes, listed before the code:
// - a formula goes to the file with its leading =, which the file format does not hold, or loses its text or its shown value on export, or a number comes back as text
// - text that starts with = comes back as a formula, or a guarded '= comes back with its apostrophe
// - a number format is lost, or a format the file format does not know is written
// - an imported sheet lands on a sheet that is there, or a taken name (any case) or a banned character passes
// - an import that fails half way leaves the sheets it already added
// - a file with two far-apart cells builds a huge grid, or the limit counts only one sheet
// - 101 sheets, a file over the size limit, or bytes that are no file pass
// - an imported cell loses the author kind of the importer, or takes a human kind by default
// - an export of a sheet over the cell limit runs
// - a formula that names a sheet reads another sheet after the sheet is renamed on import
// - an import of empty sheets takes a bad principal, or a lone = changes on the way through
// - the export limit counts one sheet at a time, or a long name is cut only when it clashes
// - a name with an upper-case or unknown extension picks a format

const human: Principal = { id: "u1", kind: "human" };
const agent: Principal = { id: "bot", kind: "agent_of_human" };
const FORMATS: SheetFormat[] = ["ods", "xlsx"];
// A new Workbook has no sheet. This one has Sheet1, as a person's workbook does.
function fresh(): Workbook {
  const book = new Workbook();
  book.addSheet("Sheet1");
  return book;
}

function sample(): Workbook {
  const book = fresh();
  book.writeRange("Sheet1", "A1", [[1, "007", true, "=A1+1", "'=not a formula"]], human);
  book.setFormat("Sheet1", "A1", { numberFormat: NUMBER_FORMATS[1] }, human);
  return book;
}

describe.each(FORMATS)("round trip through %s", (format) => {
  test("values, formulas, text, and number formats survive", () => {
    const book = sample();
    const { bytes, warnings } = exportFile(format, book);
    expect(warnings).toEqual([]);
    const target = fresh();
    const { names } = importFile(format, bytes, target, agent);
    expect(names).toEqual(["Sheet1 (2)"]);
    const name = names[0]!;
    const engine = new Engine(target);
    expect(engine.values(name, "A1:E1")[0]).toEqual([1, "007", true, 2, "'=not a formula"]);
    engine.dispose();
    const cells = target.readRange(name, "A1:E1")[0]!;
    expect(cells[3]!.formula).toBe("=A1+1");
    expect(cells[4]!.formula).toBeNull();
    expect(cells[0]!.format.numberFormat).toBe(NUMBER_FORMATS[1]);
    expect(cells[2]!.value).toBe(true);
  });

  test("the author of an import is the principal that imports", () => {
    const target = fresh();
    const [name] = importFile(format, exportFile(format, sample()).bytes, target, agent).names;
    expect(target.readRange(name!, "A1")[0]![0]!.author).toMatchObject({ id: "bot", kind: "agent_of_human", machine: true });
  });

  test("a name that is taken gets a number, in any case, and the sheets that are there stay", () => {
    const target = fresh();
    target.writeRange("Sheet1", "A1", [["mine"]], human);
    target.addSheet("sheet1 (2)");
    const { names } = importFile(format, exportFile(format, sample()).bytes, target, human);
    expect(names).toEqual(["Sheet1 (3)"]);
    expect(target.sheetNames()).toEqual(["Sheet1", "sheet1 (2)", "Sheet1 (3)"]);
    expect(target.readRange("Sheet1", "A1")[0]![0]!.value).toBe("mine");
  });
});

describe("the neutral workbook", () => {
  test("a formula cell carries its text and its shown value, and an error value is left out", () => {
    const book = fresh();
    book.writeRange("Sheet1", "A1", [[2, "=A1*3", "=1/0", "=A1&\"x\""]], human);
    const { cells } = toFileBook(book).sheets[0]!;
    expect(cells.B1).toMatchObject({ formula: "A1*3", value: 6 });
    expect(cells.C1).toEqual({ formula: "1/0" });
    expect(cells.D1).toMatchObject({ value: "2x" });
  });

  test("a format that agentic-formats does not know is not written", () => {
    const book = fresh();
    book.writeRange("Sheet1", "A1", [[1]], human);
    book.setFormat("Sheet1", "A1", { numberFormat: "[Red]0.000000000" }, human);
    expect(toFileBook(book).sheets[0]!.cells.A1).toEqual({ value: 1 });
  });

  test("guarded text leaves the workbook without its apostrophe, and a lone = stays text", () => {
    const book = fresh();
    book.writeRange("Sheet1", "A1", [["'=x", "="]], human);
    expect(toFileBook(book).sheets[0]!.cells).toEqual({ A1: { value: "=x" }, B1: { value: "=" } });
    const target = fresh();
    const [name] = importFile("xlsx", exportFile("xlsx", book).bytes, target, human).names;
    expect(target.readRange(name!, "A1:B1")[0]!.map((c) => c.value)).toEqual(["'=x", "="]);
  });

  test("an empty sheet exports as a sheet with no cells", () => {
    expect(toFileBook(fresh()).sheets).toEqual([{ name: "Sheet1", cells: {} }]);
  });
});

describe("limits and bad input", () => {
  const far: FileBook = { sheets: [{ name: "Far", cells: { A1: { value: 1 }, XFD1048576: { value: 2 } } }] };

  test("two far-apart cells are refused, fast, and nothing is added", () => {
    const target = fresh();
    const start = Date.now();
    expect(() => addSheets(target, far, human)).toThrow(/too large to import/);
    expect(Date.now() - start).toBeLessThan(2_000);
    expect(target.sheetNames()).toEqual(["Sheet1"]);
  });

  test("a failure on a later sheet removes the sheets that were added before it", () => {
    const target = fresh();
    const ok: FileBook["sheets"][number] = { name: "Fine", cells: { A1: { value: "kept?" } } };
    expect(() => addSheets(target, { sheets: [ok, ...far.sheets] }, human)).toThrow(/too large to import/);
    expect(target.sheetNames()).toEqual(["Sheet1"]);
  });

  test("the limit counts the grids of all the sheets", () => {
    // Each grid is 26 x 20,000 = 520,000 cells: one fits under the limit of 1,000,000, and two do not.
    const wide = (name: string): FileBook["sheets"][number] => ({ name, cells: { A1: { value: 1 }, Z20000: { value: 2 } } });
    const one = fresh();
    expect(addSheets(one, { sheets: [wide("W0")] }, human)).toEqual(["W0"]);
    const two = fresh();
    expect(() => addSheets(two, { sheets: [wide("W0"), wide("W1")] }, human)).toThrow(/too large to import/);
    expect(two.sheetNames()).toEqual(["Sheet1"]);
  });

  test("a formula that names a sheet that is renamed is refused, and one that is not renamed is kept", () => {
    const file: FileBook = { sheets: [{ name: "Summary", cells: { A1: { formula: "Data!A1" } } }, { name: "Data", cells: { A1: { value: "NEW" } } }] };
    const taken = fresh();
    taken.addSheet("Data");
    taken.writeRange("Data", "A1", [["OLD"]], human);
    expect(() => addSheets(taken, file, human)).toThrow(/would be renamed "Data \(2\)", and a formula in the file names a sheet/);
    expect(taken.sheetNames()).toEqual(["Sheet1", "Data"]);
    const clean = fresh();
    addSheets(clean, file, human);
    const engine = new Engine(clean);
    expect(engine.values("Summary", "A1")[0]![0]).toBe("NEW");
    engine.dispose();
  });

  test("a sheet name with an apostrophe is renamed, so a formula that quotes it is refused", () => {
    const file: FileBook = { sheets: [{ name: "S", cells: { A1: { formula: "'Bob''s'!A1" } } }, { name: "Bob's", cells: { A1: { value: 5 } } }] };
    const target = fresh();
    expect(() => addSheets(target, file, human)).toThrow(/would be renamed "Bob_s"/);
    expect(target.sheetNames()).toEqual(["Sheet1"]);
  });

  test("a rename with no formula that names a sheet goes through, even when a formula holds the letters of the name", () => {
    const file: FileBook = { sheets: [{ name: "S:1", cells: { A1: { formula: "SUM(B1:B2)" }, B1: { value: 2 } } }] };
    const target = fresh();
    expect(addSheets(target, file, human)).toEqual(["S_1"]);
  });

  test("a bad principal is refused even when every sheet is empty", () => {
    const target = fresh();
    const empty: FileBook = { sheets: [{ name: "E", cells: {} }] };
    expect(() => addSheets(target, empty, { id: "", kind: "human" })).toThrow(/invalid principal/);
    expect(() => addSheets(target, empty, { id: "x", kind: "nope" as never })).toThrow(/invalid principal/);
    expect(target.sheetNames()).toEqual(["Sheet1"]);
  });

  test("a long name is cut to 31 characters, with or without a clash", () => {
    const long = "L".repeat(40);
    const target = fresh();
    expect(addSheets(target, { sheets: [{ name: long, cells: {} }] }, human)).toEqual(["L".repeat(31)]);
    expect(addSheets(target, { sheets: [{ name: long, cells: {} }] }, human)).toEqual([`${"L".repeat(27)} (2)`]);
  });

  test(`more than ${MAX_IMPORT_SHEETS} sheets are refused`, () => {
    const sheets = Array.from({ length: MAX_IMPORT_SHEETS + 1 }, (_, i) => ({ name: `S${i}`, cells: {} }));
    const target = fresh();
    expect(() => addSheets(target, { sheets }, human)).toThrow(/up to 100 sheets/);
    expect(target.sheetNames()).toEqual(["Sheet1"]);
  });

  test.each(FORMATS)("bytes that are not a %s file throw a FormatError and add nothing", (format) => {
    const target = fresh();
    expect(() => importFile(format, new TextEncoder().encode("not a zip"), target, human)).toThrow(FormatError);
    expect(target.sheetNames()).toEqual(["Sheet1"]);
  });

  test("a file over the size limit is refused before it is read", () => {
    expect(() => importFile("xlsx", new Uint8Array(26 * 1024 * 1024), new Workbook(), human)).toThrow(/files can be up to/);
  });

  test("a sheet whose extent is over the cell limit is not exported", () => {
    const book = fresh();
    book.writeRange("Sheet1", "A1", [[1]], human);
    book.writeRange("Sheet1", "ZZ2000", [[1]], human);
    expect(() => exportFile("ods", book)).toThrow(/sheet Sheet1 is too large to export/);
  });

  test("the export limit counts the used ranges of all the sheets", () => {
    const book = fresh();
    // Each used range is 26 x 20,000 = 520,000 cells: one fits, and two do not.
    for (const name of ["Sheet1", "Two"]) {
      if (name !== "Sheet1") book.addSheet(name);
      book.writeRange(name, "Z20000", [[1]], human);
    }
    expect(() => exportFile("xlsx", book)).toThrow(/sheet Two is too large to export/);
  });
});

describe("file kinds", () => {
  test.each([
    ["a.ods", "ods"],
    ["B.XLSX", "xlsx"],
    ["a.tar.xlsx", "xlsx"],
    ["a.csv", null],
    ["xlsx", null],
    ["", null],
  ])("formatOf(%s) is %s", (name, expected) => expect(formatOf(name)).toBe(expected));

  test("each format has a MIME type", () => {
    expect(MIME.ods).toBe("application/vnd.oasis.opendocument.spreadsheet");
    expect(MIME.xlsx).toContain("spreadsheetml.sheet");
  });
});

