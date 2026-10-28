---
name: agentic-spreadsheet
description: Spreadsheets with formulas that people and agents edit together. Keywords: spreadsheet,sheets,formulas,grid,collaboration,agents,mcp,cli,csv,xlsx.
---

# agentic-spreadsheet

Spreadsheets with formulas that people and agents edit together.

## Use the SDK

```js
import { summarizeText, dispatchText } from "agentic-spreadsheet";

summarizeText("some text");
await dispatchText("some text", { dryRun: true });
```

## Use the CLI

```sh
agentic-spreadsheet summary --text "some text"
agentic-spreadsheet dispatch --text "some text" --dry-run
```

## Use the MCP server

Start the server with `agentic-spreadsheet mcp`. Call `summarize_text`
for stats, `dispatch_text` for routes, `get_cost_report` for costs.

## Rules

- Preview each route with dry run first.
- Report guard denials and adapter errors as failures.
- Read the cost report at the end.
