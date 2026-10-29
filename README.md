# agentic-spreadsheet

[![npm version](https://img.shields.io/npm/v/agentic-spreadsheet)](https://www.npmjs.com/package/agentic-spreadsheet)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://opensource.org/license/mit)

[Package](https://www.npmjs.com/package/agentic-spreadsheet) · [Source](https://github.com/pooriaarab/agentic-spreadsheet) · [Support](https://github.com/pooriaarab/agentic-spreadsheet/issues)

Spreadsheets with formulas that people and agents edit together.

> Status: early. The workbook model and the formula engine work. It is not on npm yet.
> The CLI, MCP server, and HTTP API are planned.

## Install

```sh
npm install agentic-spreadsheet
```

## Quick start

```ts
import { Engine, Workbook } from "agentic-spreadsheet";

const me = { id: "user-1", kind: "human" } as const;
const wb = new Workbook();
wb.addSheet("Sheet1");
wb.writeRange("Sheet1", "A1", [[1], [2], ["=SUM(A1:A2)"]], me);
console.log(new Engine(wb).values("Sheet1", "A1:A3")); // [[1], [2], [3]]
```

## Features

- A workbook is one Yjs document. Two copies merge. Content, each format field, and the editor stamp
  of a cell are separate keys, so two writers on one new cell keep both edits.
  If two writers set the same field, the later write wins.
- Many sheets in a set order. Add, rename, and delete sheets.
- Sparse cells keyed by A1 coordinates. Each cell holds a value or a formula, plus a format.
  Text is at most 32,767 characters. A formula is at most 8,192.
- Column widths and frozen rows.
- Insert and delete rows and columns. Formula refs follow the moved cells.
  A ref to a deleted cell becomes `#REF!`. A range shrinks or grows to match.
  An insert that would push a cell off the grid is refused. A count is at most 10,000.
- A formula engine reads the workbook. It has `+ - * / ^ & %`, comparisons, and cross-sheet refs.
  It returns `#REF!`, `#DIV/0!`, `#NAME?`, `#VALUE!`, `#NUM!`, `#N/A`, `#CIRCULAR`, and `#DEPTH!`.
- Functions: `IF`, `IFS`, `IFERROR`, `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `COUNTIF(S)`,
  `SUMIF(S)`, `VLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `AND`, `OR`, `NOT`, `ROUND`, `ROUNDUP`,
  `ROUNDDOWN`, `ABS`, `CONCAT`, `TEXT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `TRIM`,
  `DATE`, `TODAY`, `NOW`. Other names come from `@formulajs/formulajs` and are untested here.
  Blank cells in a range are skipped, as in a spreadsheet app. `MUNIT` and `EXPAND` return `#NAME?`.
- Read a range. Write a range of values or formulas.
- Every cell keeps the principal that last edited it. The kind is one of
  `human`, `agent_of_human`, `agent_of_agent`, `agent`, or `user_token`.
  Only `human` reads as human-made. All other kinds read as machine-made.

## Limits

- The engine keeps no cache between calls, so there is no incremental recalc yet.
  A chain of more than 400 dependent cells, a formula nested more than 100 levels deep, and a formula
  with more than 1,000 chained operators give `#DEPTH!` in that cell. Other cells are not affected.
  A loop gives `#CIRCULAR` in every cell of the loop, whichever cell you read first.
- Ref rewriting does not handle whole-column refs such as `A:A`.
- Insert, delete, and rename need one writer at a time. They move cells and rewrite formulas,
  and that does not merge. If two clients do it at once, `structureConflict()` turns true after the
  merge and the workbook refuses writes until someone calls `acceptStructureMerge()`.
  A value write that races an insert lands at the old address, and nothing flags it.
- There is no storage, CLI, MCP server, or HTTP API yet.
- Two clients that add a sheet with the same name while offline both succeed.
  The merged workbook then holds two sheets with one name.

## License

MIT.
