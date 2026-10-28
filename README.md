# agentic-spreadsheet

[![npm version](https://img.shields.io/npm/v/agentic-spreadsheet)](https://www.npmjs.com/package/agentic-spreadsheet)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://opensource.org/license/mit)

[Package](https://www.npmjs.com/package/agentic-spreadsheet) · [Source](https://github.com/pooriaarab/agentic-spreadsheet) · [Support](https://github.com/pooriaarab/agentic-spreadsheet/issues)

Spreadsheets with formulas that people and agents edit together.

> Status: early. Only the workbook model works today. It is not on npm yet.
> The formula engine, CLI, MCP server, and HTTP API are planned.

## Install

```sh
npm install agentic-spreadsheet
```

## Quick start

```ts
import { Workbook } from "agentic-spreadsheet";

const me = { id: "user-1", kind: "human" } as const;
const wb = new Workbook();
wb.addSheet("Sheet1");
wb.writeRange("Sheet1", "A1", [[1], [2], ["=SUM(A1:A2)"]], me);
console.log(wb.readRange("Sheet1", "A1:A4"));
```

## Features

- A workbook is one Yjs document. Two copies merge without conflict.
- Many sheets in a set order.
- Sparse cells keyed by A1 coordinates. Each cell holds a value or a formula, plus a format.
- Column widths and frozen rows.
- Read a range. Write a range of values or formulas.
- Every cell keeps the principal that last edited it. The kind is one of
  `human`, `agent_of_human`, `agent_of_agent`, `agent`, or `user_token`.
  Only `human` reads as human-made. All other kinds read as machine-made.

## Limits

- Formulas are stored as text. Nothing evaluates them yet.
- There is no storage, CLI, MCP server, or HTTP API yet.

## License

MIT.
