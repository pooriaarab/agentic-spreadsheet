# agentic-spreadsheet agent guide

Spreadsheets with formulas that people and agents edit together.

## Mental model

A workbook is a Yjs document. It holds ordered sheets. Each sheet holds
sparse cells keyed by the ids of their row and column, and an order of those ids.
An A1 address is a view of that order. A formula stores refs as ids and shows them as A1.
Code in `src/core/` never reads a cell by A1 text: it asks the sheet for the id at a position.
A cell holds a value or a formula.

## Faces

- SDK: `import { Workbook } from "agentic-spreadsheet"`.
- CLI and MCP: `npx agentic-spreadsheet <command>` and `npx agentic-spreadsheet mcp`. See the README.

## Rules

- Pass your own principal on every write: `--as` and `--kind` (or `SHEET_AS` and `SHEET_KIND`). Use an agent kind, never `human`. There is no default.
- A string that starts with `=` is a formula. Pass a number as a number.
- Read the result after you insert or delete rows or columns. Refs to deleted cells read `#REF!`.
- Report errors as errors. Never present a failed call as success.
