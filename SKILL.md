---
name: agentic-spreadsheet
description: Spreadsheets with formulas that people and agents edit together. Keywords: spreadsheet,sheets,formulas,grid,collaboration,agents,mcp,cli,csv,xlsx.
---

# agentic-spreadsheet

Spreadsheets with formulas that people and agents edit together.

## Use the SDK

```js
import { Workbook } from "agentic-spreadsheet";

const agent = { id: "agent-1", kind: "agent_of_human" };
const wb = new Workbook();
wb.addSheet("Sheet1");
wb.writeRange("Sheet1", "A1", [[1, 2, "=A1+B1"]], agent);
wb.readRange("Sheet1", "A1:C1");
```
