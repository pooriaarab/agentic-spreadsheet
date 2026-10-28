# agentic-spreadsheet agent guide

Spreadsheets with formulas that people and agents edit together.

## Mental model

A workbook is a Yjs document. It holds ordered sheets. Each sheet holds
sparse cells keyed by A1 coordinates. A cell holds a value or a formula.

## Faces

- SDK: `import { Workbook } from "agentic-spreadsheet"`.
- CLI and MCP: planned. Only `--help` and `--version` work today.

## Rules

- Pass your own principal on every write. Use an agent kind, never `human`.
- A string that starts with `=` is a formula. Pass a number as a number.
- Read the result after you insert or delete rows or columns. Refs to deleted cells read `#REF!`.
- Report errors as errors. Never present a failed call as success.
