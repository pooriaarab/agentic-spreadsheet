# agentic-spreadsheet

[![npm version](https://img.shields.io/npm/v/agentic-spreadsheet)](https://www.npmjs.com/package/agentic-spreadsheet)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://opensource.org/license/mit)

[Package](https://www.npmjs.com/package/agentic-spreadsheet) · [Source](https://github.com/pooriaarab/agentic-spreadsheet) · [Support](https://github.com/pooriaarab/agentic-spreadsheet/issues)

Spreadsheets with formulas that people and agents edit together.

> Status: early. The workbook model, formula engine, CLI, and MCP server work. The package is on npm.
> The HTTP API is a stopgap with in-memory storage.
>
> Warning: the HTTP API (the Worker) refuses to serve until you set `SHEETS_API_KEYS`, or set
> `SHEETS_UNSAFE_TRUST_HEADERS=1`. That second setting trusts any caller who sets two headers.
> Use it for tests only. See [HTTP API](#http-api).
>
> Warning: the CLI and the MCP server do not check who calls them. `--as` and `--kind` are claims
> that label each edit. Do not give either one to a caller you do not trust.

## Install

```sh
npm install agentic-spreadsheet
```

Node 22 or later.

## Quick start

```ts
import { Engine, Workbook } from "agentic-spreadsheet";

const me = { id: "user-1", kind: "human" } as const;
const wb = new Workbook();
wb.addSheet("Sheet1");
wb.writeRange("Sheet1", "A1", [[1], [2], ["=SUM(A1:A2)"]], me);
console.log(new Engine(wb).values("Sheet1", "A1:A3")); // [[1], [2], [3]]
```

## CLI and MCP server

```sh
export SHEET_DIR=./.sheets          # where workbooks live. The default is ./.sheets
export SHEET_AS=ana SHEET_KIND=human   # who writes. A write fails without both
npx agentic-spreadsheet new budget
npx agentic-spreadsheet write budget A1 '[[10,20,"=A1+B1"]]'
npx agentic-spreadsheet read budget A1:C1   # values, formulas, and authors
npx agentic-spreadsheet mcp                 # MCP server on stdio
```

The CLI and the MCP server share one list of operations: `new`, `list`, `read`, `write`,
`formula`, `insert-rows`, `export` (CSV), `import`. MCP tool names start with `sheet_`.
A write needs `--as <id>` and `--kind <kind>` (or `SHEET_AS` and `SHEET_KIND`). There is no default
kind. An agent sets `agent`, `agent_of_human`, or `agent_of_agent`. A read needs neither.
`import` takes `--as-text`: a cell that starts with `=` stays text. Use it for CSV you do not trust.
`export` and `import` handle sheets of 150,000 rows and more. Export puts a `'` before text that
starts with `+ - @` or a control character, and import removes it, so a round trip keeps the text.

## HTTP API

The Worker serves the same operations under `/api/v1/sheets`. Run it with `npx wrangler dev`.

```sh
# one key, at least 32 bytes, tied to one id and kind
npx wrangler dev --var 'SHEETS_API_KEYS:{"<your-key>":{"id":"bot-1","kind":"agent"}}'
curl -H "authorization: Bearer <your-key>" -X POST localhost:8787/api/v1/sheets -d '{"book":"budget"}'
curl -H "authorization: Bearer <your-key>" -X PUT localhost:8787/api/v1/sheets/budget/range \
  -d '{"anchor":"A1","values":[[10,20,"=A1+B1"]]}'
curl -H "authorization: Bearer <your-key>" 'localhost:8787/api/v1/sheets/budget/range?range=A1:C1'
```

| Method and path | Operation |
|---|---|
| `GET /` and `POST /` | `list` and `new` |
| `GET /:book/range` and `PUT /:book/range` | `read` and `write` |
| `PUT /:book/formula` | `formula` |
| `POST /:book/rows` | `insert-rows` |
| `GET /:book/export` and `POST /:book/import` | `export` (CSV) and `import` |
| `PUT /:book/share` and `DELETE /:book/share/:kind/:subject` | Give a person (`human`) or an agent (`agent`) a role (`viewer`, `commenter`, `editor`), and take it back |
| `DELETE /:book` | The owner deletes the workbook, and its grants, invites, links, and audit entries go with it |

The key sets who writes, so a body cannot name another principal. A key must have 32 bytes or more,
and each needs an id and a kind. A missing or bad key gets 401. Errors are 400, 403, 404, 409, or 413.
A request body is at most 5 MB. `GET /health` needs no key.

A workbook has an owner: the principal that made it. The owner shares it with a kind (`human` or `agent`), an id, and a role.
A viewer or commenter reads. An editor reads and writes. Only the owner shares. `agentic-identity` keeps the
roles and writes an audit entry for each check. A principal with no role gets 404, as for a book that does not
exist. A principal with a role that is too low gets 403. A share follows the kind and the id: a person and an agent with one id hold separate roles.
Every kind but `human` counts as an agent.
The owner cannot lower its own role. Each check adds an audit entry to the store, and nothing limits how many.
If the Worker fails after it saves a new workbook and before it records the owner, the book has no owner.

## Features

- A workbook is one Yjs document. Two copies merge. Content, each format field, and the editor stamp
  of a cell are separate keys, so two writers on one new cell keep both edits.
  If two writers set the same field, the later write wins.
- Many sheets in a set order. Add, rename, and delete sheets.
- Sparse cells. Each cell holds a value or a formula, plus a format.
  Text is at most 32,767 characters. A formula is at most 8,192.
- Every row and column has an id that never changes. A cell is keyed by the ids of its row and its
  column, and a formula keeps its refs as ids. An A1 address is the position of those ids, so the SDK,
  the CLI, the MCP server, and the HTTP API take and return A1. Insert, delete, and move change the
  order of the ids and nothing else. Two writers who edit one sheet at the same time lose nothing,
  and a write lands on the row it was meant for, whatever the other writer inserts, deletes, or moves.
  `moveRows` and `moveCols` are in the SDK.
- A workbook saved by an older version, with A1 keys, converts when it loads.
- Column widths and frozen rows.
- Insert and delete rows and columns. Formula refs follow the moved cells.
  A ref to a deleted cell becomes `#REF!`. A range shrinks or grows to match.
  An insert that would push a cell off the grid is refused. A count is at most 10,000.
- A formula engine reads the workbook. It has `+ - * / ^ & %`, comparisons, and cross-sheet refs.
  It keeps values between reads. A change recomputes only the cells that read it, and a chain of
  100,000 cells works. Call `engine.dispose()` when the workbook outlives the engine.
  It returns `#REF!`, `#DIV/0!`, `#NAME?`, `#VALUE!`, `#NUM!`, `#N/A`, `#CIRCULAR`, and `#DEPTH!`.
- Functions: `IF`, `IFS`, `IFERROR`, `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `COUNTIF(S)`,
  `SUMIF(S)`, `VLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `AND`, `OR`, `NOT`, `ROUND`, `ROUNDUP`,
  `ROUNDDOWN`, `ABS`, `CONCAT`, `TEXT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `TRIM`,
  `DATE`, `TODAY`, `NOW`. Other names come from `@formulajs/formulajs` and are untested here.
  Blank cells in a range are skipped, as in a spreadsheet app. `MUNIT` and `EXPAND` return `#NAME?`.
- CSV import and export. Export marks text that starts with `= + - @` so it cannot run as a formula.
- Read a range. Write a range of values or formulas.
- Every cell keeps the principal that last edited it. The kind is one of
  `human`, `agent_of_human`, `agent_of_agent`, `agent`, or `user_token`.
  Only `human` reads as human-made. All other kinds read as machine-made.

## Limits

- A formula nested more than 100 levels deep, or with more than 1,000 chained operators, gives
  `#DEPTH!` in that cell. Other cells are not affected. A chain of cells has no depth limit.
- A loop gives `#CIRCULAR` in every cell on the loop, whichever cell you read first.
- A change to a cell finds the formulas that read it through a range by a scan of all ranges on
  that sheet. A sheet with very many range formulas pays for that scan on each edit.
- Ref rewriting does not handle whole-column refs such as `A:A`.
- A write to a row that another writer deletes goes with the row. Two writers who move one row at
  once get one place for it, the same on both copies.
- Two copies that each convert an old A1-keyed workbook, and then merge, are flagged:
  `conversionClash()` is true and every read and write throws, because each copy made its own order
  for the sheets and the merge keeps one. Load the workbook again from one copy, or better, convert
  once, save it, and then share it. The first write of the CLI or the Worker converts a stored
  workbook and saves it in one committed write, so a stored workbook converts once. A read converts a
  copy in memory and saves nothing. A copy of one old file converted on two machines meets this case.
  An old workbook whose log shows row, column, or sheet-name edits made at the same time and not
  accepted does not convert: the load throws, because its cells may sit at the wrong address. Accept the
  merge in the version that saved it, then load it again. `structureConflict()` and
  `acceptStructureMerge()` are gone from the SDK, because inserts, deletes, and moves now merge.
  A clash that was saved cannot be repaired in place. Restore the workbook from a copy saved before it.
  Old A1-keyed state that a sync provider delivers after the workbook is made is refused with an
  error. Load it with `Workbook.fromState` instead.
  A sheet that comes without an id after a conversion converts on the next `Workbook.fromState`. It is a
  clash only when two copies convert that same sheet.
  A sheet that has an id but holds an old A1 cell key is refused with an error, because that cell would
  read as empty. Restore the workbook from a copy saved before that key arrived. An old formula that
  holds U+0001 or U+0002 does not convert: the load throws. Remove the character in the version that
  saved it, then load the workbook again.
- A single ref follows a moved row. A range keeps the rows that stay between its ends: a row that
  moves out of a range leaves it, and a row that moves in joins it.
- Two inserts made at once can push a cell past the last row. The workbook keeps the cell, but no
  read, export, or count sees it until rows are deleted.
- A formula that names a sheet that does not exist yet takes the sheet when it is added. A formula
  written by one writer while another adds that sheet is not bound until the next rename or add.
- Two clients that add a sheet with the same name while offline both succeed.
  The merged workbook then holds two sheets with one name.
- The Worker keeps workbooks in memory, in one isolate. A restart loses them, and two isolates do not
  share them. A workbook stored before sharing has no owner,
  and the Worker refuses it for every principal. `list` still names every workbook to every valid key.
  There is no rate limit yet.
- Storage uses the `Store` and `Blobs` ports of `agentic-identity`. A workbook is a document in the store
  and its state is a blob. The CLI keeps both in files under the store directory (`store.json` and
  `blobs/`). A write takes a lock file, and a lock from a dead process is taken over. A workbook saved by
  0.1.0 as one file per workbook is read as it is and moves into `store.json` on the first write. A 0.1.0 file whose id starts and ends with two underscores,
  or is longer than 256 characters, stops the store with an error that names the file. Rename it first.
  XLSX and ODS import and export are not built yet.
- CSV export refuses a sheet whose extent passes 1,000,000 cells. The HTTP API refuses a write or import of
  more than 10,000 cells (413). A formula may read at most 1,000,000 range cells.

## License

MIT.
