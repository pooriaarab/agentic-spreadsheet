# agentic-spreadsheet

[![npm version](https://img.shields.io/npm/v/agentic-spreadsheet)](https://www.npmjs.com/package/agentic-spreadsheet)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://opensource.org/license/mit)

[Package](https://www.npmjs.com/package/agentic-spreadsheet) · [Source](https://github.com/pooriaarab/agentic-spreadsheet) · [Support](https://github.com/pooriaarab/agentic-spreadsheet/issues)

Spreadsheets with formulas that people and agents edit together.

> Status: early. The workbook model, formula engine, CLI, and MCP server work. The package is on npm.
> The HTTP API keeps data in D1 and R2. Keys are read through the `Auth` port of `agentic-identity`.
>
> Warning: the HTTP API (the Worker) refuses to serve until you set `SHEETS_API_KEYS`, or set
> `SHEETS_UNSAFE_TRUST_HEADERS=1`. That second setting trusts any caller who sets two headers.
> Use it for tests only. See [HTTP API](#http-api).
>
> Warning: the Worker also refuses to serve until `SHEETS_REGION` says where the data lives (`eu` or `us`).
> Under `wrangler dev`, set `SHEETS_REGION=unknown` and `SHEETS_UNSAFE_REGION=1`.
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

## SDK

`agentic-spreadsheet/client` calls the Worker over HTTP. It uses only `fetch`, so it runs in browsers,
Workers, and Node 22. It has one method for each route.

```ts
import { createClient, SheetsApiError } from "agentic-spreadsheet/client";

const sheets = createClient({ baseUrl: "https://sheets.example.com", token: "<your-key>" });
await sheets.create({ book: "budget" });
await sheets.write({ book: "budget", anchor: "A1", values: [[10, 20, "=A1+B1"]] });
const { values } = await sheets.read({ book: "budget", range: "A1:C1" }); // [[10, 20, 30]]

try {
  await sheets.read({ book: "missing", range: "A1" });
} catch (error) {
  if (error instanceof SheetsApiError) console.log(error.status, error.code, error.message); // 404 "not_found" ...
}
```

Methods: `list`, `create`, `read`, `write`, `setFormula`, `insertRows`, `deleteRows`, `insertCols`,
`deleteCols`, `sheets`, `addSheet`, `deleteSheet`, `exportCsv`, `importCsv`, `share`, `unshare`,
`remove`, and `health`. Each takes the same fields as its route, plus an optional
`{ signal }` as the second argument to cancel the call. Pass `fetch` to `createClient` to use your own.

A failed call throws `SheetsApiError` with the HTTP `status` and a `code`: `bad_request` (400),
`unauthorized` (401), `forbidden` (403), `not_found` (404), `conflict` (409), `too_large` (413), or
`server_error` (5xx). Any other status, such as a 429 from a proxy, is `unknown`. The API sends no code of its own, so the client names the status. A network failure, an abort, or
a 2xx answer that is not JSON (for example from a wrong `baseUrl`) throws the error that `fetch` or `json()` throws. The client never retries. A write that fails may have landed:
read the book before you send it again.

### ODS and XLSX

`agentic-spreadsheet/formats` maps a workbook to and from ODS and XLSX, through `agentic-formats`.

```ts
import { exportFile, importFile, formatOf, MIME } from "agentic-spreadsheet/formats";

const { bytes, warnings } = exportFile("xlsx", book); // MIME.xlsx
const { names } = importFile("ods", odsBytes, book, { id: "bot", kind: "agent_of_human" });
```

- `exportFile(format, book)` returns the bytes and the notes of what the format could not hold. A formula cell
  keeps its formula and its shown value. The used ranges of all the sheets together can hold up to 1,000,000 cells, or the call throws a `SheetError`.
- `importFile(format, bytes, book, by)` adds the sheets of the file after the sheets that are there. A taken
  name gets a number, and a name is cut to 31 characters. A file that needs a rename, and has a formula that names a sheet (a
  `!` in the formula), throws a `SheetError`, because the formula would point at another sheet. A formula that
  names a sheet that is not in the file binds to the sheet of that name in the workbook, if there is one. Every imported cell has the author `by`. All of the sheets are added, or none. A file over
  25 MB, over 100 sheets, or over 1,000,000 cells in the grids of its sheets throws a `SheetError`. Bytes that are
  not a file of that format throw the `FormatError` of `agentic-formats`.
- `formatOf(fileName)` returns `"ods"`, `"xlsx"`, or `null`. `toFileBook(book)` and `addSheets(book, model, by)`
  work with the neutral model of `agentic-formats`.

### Other exports

The root export, `agentic-spreadsheet`, also exports these names.

- A1 helpers. `parseCell("B2")` returns `{ col: 2, row: 2 }`, and `formatCell` turns it back into `"B2"`.
  `parseRange("A1:C3")` returns `{ c1, r1, c2, r2 }`. `colIndex("AA")` returns `27`, and `colLetters(27)`
  returns `"AA"`. `parseCell` and `parseRange` throw on bad text.
- CSV helpers. `parseCsv(text)` returns rows of strings. `formatCsv(rows)` writes CRLF rows. It adds a leading
  `'` to text that starts with `= + - @`, a tab, or a carriage return. `inferCell(text)` turns one field into
  a number, a boolean, a formula, or text. Empty text gives `null`. With `asText` set to `true`, a leading `=`
  is kept as text, with a `'` in front.
- Stores. `createMemoryStore()` and `createMemoryBlobs()` keep data in memory, for tests. `createFileStore(dir)`
  and `createFileBlobs(dir)` keep data in files, as the CLI does. A store runs a transaction again when another
  writer changed what it read, and throws `StoreConflictError` when the transaction keeps losing.
  `createFileStore` also throws it when it keeps losing to another process, or when the file lock stays held. Read the book,
  then retry the write.
- `ctxFromEnv(process.env)` builds the context that the CLI and the MCP server use. It reads `SHEET_DIR`,
  `SHEET_AS`, and `SHEET_KIND`.
- `OPS` lists the operations that the CLI, the MCP server, and the HTTP API run. Each has a `name`, a
  `description`, the `args` names, a `shape` (a Zod shape), and a `run` function. Parse the arguments with
  `z.object(op.shape)` before you call `run`. `OpError` carries the HTTP status (404, 409, or 413).
- `SheetError` is the error that a `Workbook` method throws for bad input, such as a bad sheet name, column,
  width, position, value, or principal.
  A bad A1 address throws a plain `Error`.
- `toBase64(bytes)` and `fromBase64(text)` convert between bytes and base64.
- `PRINCIPAL_KINDS` lists the five principal kinds. `isMachineMade(kind)` returns `true` for every kind except
  `human`.
- `VERSION` is the package version.

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
`formula`, `insert-rows`, `delete-rows`, `insert-cols`, `delete-cols`, `sheets`, `add-sheet`, `delete-sheet`,
`export` (CSV), `import`. MCP tool names start with `sheet_`.

To share a workbook, use a Worker. A local store has no roles, so these commands need `SHEET_URL` and the key of the owner in `SHEET_TOKEN`. Use an `https` URL: the key goes in a header.

```sh
export SHEET_URL=https://sheets.example.com SHEET_TOKEN=<key>
npx agentic-spreadsheet share budget --subject ana --subject-kind human --role editor   # give a role (owner only)
npx agentic-spreadsheet shares budget                                                   # list the roles (owner only)
npx agentic-spreadsheet unshare budget --subject ana --subject-kind human               # take it away
```

The role is `viewer`, `commenter`, or `editor`. `--kind` is the kind of a writer, so a share uses `--subject-kind`.
`shares` shows each subject as a hash of its id, because the store never keeps the id. A refused call prints
`error: <message>` and exits 1. The Worker has no share links: share with named people and agents. With the same
two variables, `mcp` also has `sheet_share`, `sheet_shares`, and `sheet_unshare`. Without them, these tools do not exist.
A write needs `--as <id>` and `--kind <kind>` (or `SHEET_AS` and `SHEET_KIND`). There is no default
kind. An agent sets `agent`, `agent_of_human`, or `agent_of_agent`. A read needs neither.
`import` takes `--as-text`: a cell that starts with `=` stays text. Use it for CSV you do not trust.
`export` and `import` handle sheets of 150,000 rows and more. Export puts a `'` before text that
starts with `+ - @` or a control character, and import removes it, so a round trip keeps the text.

## Live demo

A demo Worker runs at <https://agentic-spreadsheet-demo.hello-fd1.workers.dev>. It is a staging build. Open the
address to see the web UI. `GET /health` answers 200 with no key. The API needs a key, and the demo does not
publish one, so run your own Worker to try writes (see [HTTP API](#http-api)). The demo may be reset at any time.

The demo uses `wrangler.demo.jsonc`: a D1 database in `weur`, an R2 bucket with the `eu` jurisdiction, and
`SHEETS_REGION=eu`. To redeploy, a maintainer with a Cloudflare token runs:

```sh
CLOUDFLARE_API_TOKEN=<token> npm run deploy:demo
```

The API keys live in a Worker secret, never in the repo or in CI. Set them with
`npx wrangler secret put SHEETS_API_KEYS --config wrangler.demo.jsonc`. To make your own demo, copy
`wrangler.demo.jsonc`, change the name, and put your own database id and bucket name in it.

## HTTP API

The Worker serves the same operations under `/api/v1/sheets`. Run `npm run build` first, because `wrangler.jsonc`
serves `dist/ui` and wrangler needs that directory. Then run `npx wrangler dev`.

Workbooks live in a D1 database (binding `DB`) and an R2 bucket (binding `BLOBS`), both declared in
`wrangler.jsonc`. They survive a restart. `wrangler dev` keeps them in `.wrangler/state`. To deploy, create the
database and the bucket (commands are in `wrangler.jsonc`), put the database id in the file, and set
`SHEETS_REGION` to `eu` or `us`. The Worker answers 500 if a binding or the region is missing, and that includes
`/health`.

The shipped `wrangler.jsonc` is for the EU: a bucket with the `eu` jurisdiction, and a database made with
`--location weur`. For `us`, create the database with a US location, drop the jurisdiction from the bucket, and set
`SHEETS_REGION=us`. D1 reports where it served each query, and the store refuses a mismatch with the claim (the
Worker answers 500). R2 cannot report its jurisdiction, so nothing checks the bucket: keep the claim and the bucket in step.

```sh
# one key, at least 32 bytes, tied to one id and kind (a key may also carry "scopes" and, for an agent that acts for a person, "parent")
npx wrangler dev --var SHEETS_REGION:unknown --var SHEETS_UNSAFE_REGION:1 --var 'SHEETS_API_KEYS:{"<your-key>":{"id":"bot-1","kind":"agent"}}'
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
| `POST /:book/rows` and `DELETE /:book/rows?at=&count=` | `insert-rows` and `delete-rows` |
| `POST /:book/cols` and `DELETE /:book/cols?at=&count=` | `insert-cols` and `delete-cols` (`at` is a column number, A is 1) |
| `GET /:book/sheets`, `POST /:book/sheets`, and `DELETE /:book/sheets/:name` | `sheets`, `add-sheet`, and `delete-sheet` |
| `GET /:book/export` and `POST /:book/import` | `export` (CSV) and `import` |
| `GET /:book/shares` | List the roles on the workbook (owner only). A subject shows as a hash |
| `PUT /:book/share` and `DELETE /:book/share/:kind/:subject` | Give a person (`human`) or an agent (`agent`) a role (`viewer`, `commenter`, `editor`), and take it back |
| `DELETE /:book` | The owner deletes the workbook, and its grants, invites, links, and audit entries go with it |

The key sets who writes, so a body cannot name another principal. The Worker reads each request through the
`Auth` port of `agentic-identity`. On start, it writes the keys of `SHEETS_API_KEYS` into the store as principals
and credentials. The store holds the hash of a key, never the key. A key has 32 bytes or more, with only letters,
digits, `_` and `-`, at most 512 characters, and each needs an id and a kind. A kind that acts for another principal
(`agent_of_human`, `agent_of_agent`, `user_token`) needs a `parent`, the id of another entry. `human` and `agent`
take no parent. An entry can also set `scopes`, such as `["workbook:read"]`: the key then cannot write, whatever its
role on a workbook.
A key that you remove from the env stops working in that Worker at once (its hash stays in the store until you
revoke it with `agentic-identity`). Use random keys, for example the output of `openssl rand -hex 32`. Each key has its own id.
A session cookie is not a credential. A missing or bad key gets 401.
A principal with the `viewer` role (see `PUT /:book/share`) reads and gets 403 on a write.
Errors are 400, 403, 404, 409, or 413.
A request body is at most 5 MB. `GET /health` needs no key.

A workbook has an owner: the principal that made it. The owner shares it with a kind (`human` or `agent`), an id, and a role.
A viewer or commenter reads. An editor reads and writes. Only the owner shares. `agentic-identity` keeps the
roles and writes an audit entry for each check. A principal with no role gets 404, as for a book that does not
exist. A principal with a role that is too low gets 403. A share follows the kind and the id: a person and an agent with one id hold separate roles.
Every kind but `human` counts as an agent.
The owner cannot lower its own role. Each check adds an audit entry to the store, and nothing limits how many.
If the Worker fails after it saves a new workbook and before it records the owner, the book has no owner.

## Mount the sheet in your page

`agentic-spreadsheet/ui` renders the sheet into an element. It needs no framework in your app.
You must install `react` and `react-dom` (version 19) yourself: `./ui` needs them as peer dependencies.
Load the styles once with `agentic-spreadsheet/ui.css`. The file has the `agentic-design` tokens, the component styles, and the fonts
it needs. It has no rule for `*`, `html`, `body`, headings, or links, so it leaves your page alone. The grid classes start with `as-`.
The sheet puts one `<h1>` with the workbook name in your page. The sheet does not set the theme of
your page: `agentic-design` reads `data-theme` on `<html>`, and your page owns that attribute.
The sheet uses the typed client, so the page must reach the Worker API.

```ts
import { createClient } from "agentic-spreadsheet/client";
import { mount } from "agentic-spreadsheet/ui";
import "agentic-spreadsheet/ui.css";

const client = createClient({ baseUrl: "https://sheets.example.com", token });
const sheet = mount(document.getElementById("sheet")!, {
  client,
  docId: "budget",
  principal: { kind: "human", name: "Ada" }, // kind is required: there is no default
  readOnly: false, // true shows the cells and allows no edit
});

sheet.update({ client, docId: "forecast", principal: { kind: "human", name: "Ada" } }); // show another workbook
sheet.unmount(); // remove the sheet and stop its polling
```

- `principal` needs a known `kind` and a `name`. A missing or unknown kind throws a `TypeError`, and so does
  a bad `client` or an empty `docId`. The kind is only checked. It does not change what the server allows:
  the server checks every write against the key of the client. The name shows in the presence room.
- `readOnly` turns off cell edits, the formula bar, and the row, column, and sheet buttons.
- `token` is optional. With it, the sheet joins the presence room on the origin of the page, and it sends
  the key in a WebSocket subprotocol entry, not in the URL. Do not pass `token` when the page and the Worker use
  different origins. Without it, the sheet shows no peers. See "Hosting" below for what presence needs.
- `onUnauthorized` runs when the server rejects the key.
- A React app can use the `WorkbookView` component from the same entry instead. Put it inside the
  providers of `agentic-design/react`. The web UI of this repo uses it.

## UI

A small web UI comes with the Worker in this repo. The UI source is in `ui/`, and `src/ui-worker.ts` serves it.
It uses the `agentic-design` package for tokens, controls, and dialogs, the typed client (`./client`) for
every call, and a presence room from `agentic-realtime`.
It has a list of workbooks, a sheet bar, a grid with A1 cells and a formula bar, edits for values
and formulas, add and remove for rows, columns, and sheets, and light and dark themes.

Run it from a checkout:

```sh
npm ci
npm run build                      # builds the Worker and the UI into dist/
echo 'SHEETS_API_KEYS={"a-key-of-at-least-32-bytes-0123456789":{"id":"ana","kind":"human"}}' > .dev.vars
echo 'SHEETS_REGION=unknown' >> .dev.vars && echo 'SHEETS_UNSAFE_REGION=1' >> .dev.vars
WRANGLER_SEND_METRICS=false npx wrangler dev
```

Open the address that Wrangler prints, and paste the key. The tab keeps the key until you close
it. `.dev.vars` stays out of git. Data lives in a local D1 and R2 under `.wrangler/state`, so it survives a restart.

- **Keyboard.** Arrow keys move. Enter or F2 edits the cell. Typing on a cell replaces its text.
  Enter saves and moves down. Tab saves and moves right while you edit. Escape cancels. Delete clears.
  Outside an edit, Tab walks to the next cell and leaves the grid at the last cell, so focus is never trapped.
- **Agent marker.** A cell an agent wrote has a robot mark and a left bar, and its name says who wrote it.
  Color is never the only sign.
- **Presence.** Each workbook has a room. The room shows who has the workbook open, and which cell
  each person selects. Cells do not travel through the room: the API stores them, and the grid
  reads again every 4 seconds. Two people who edit one cell at once get the last write.
- **Tests.** `npm run test:e2e` runs Playwright against `wrangler dev`. It signs in, makes a sheet,
  enters a value and a formula, checks the result and a reload, checks the keys, the structure
  changes, the agent marker, and presence, and runs an axe WCAG 2.2 AA check on each view in light and dark.
  Screenshots of each view in both themes are in [`docs/screenshots`](https://github.com/pooriaarab/agentic-spreadsheet/tree/main/docs/screenshots).
  Run `UPDATE_SCREENSHOTS=1 npm run test:e2e` to make them again.
- **Limits.** The grid shows the used range plus some room, up to 52 columns and 200 rows. The browser
  sends the key in the `agentic-spreadsheet.bearer.<key>` subprotocol entry when it joins a room, because a
  browser cannot set a header on a WebSocket. The Worker does not read a `token` query parameter: a URL
  reaches logs and proxies. Use HTTPS, and keep Worker logs private. There is no sign-in page for a person
  yet: the key is the sign-in.
- **Hosting.** The npm package ships the UI files in `dist/ui`, but not the Worker that serves them.
  Read the files by path, from `node_modules/agentic-spreadsheet/dist/ui`. `exports` does not list them.
  A host must serve them at the root of an origin, because the asset paths start at `/`, and the UI calls
  the API on its own origin.
  The host answers every path under `/api/v1/sheets/`, and `/health`, with the default export of
  `agentic-spreadsheet/worker`. Pass an `env` with `SHEETS_API_KEYS`, `SHEETS_REGION`, and the `DB` and `BLOBS` bindings, or the Worker answers 500.
  Presence does not work this way yet. Its socket needs `agentic-realtime` and a `SyncRoom` Durable Object,
  and `roomGrant` needs the same `DB` binding. For presence, run the Worker from
  a checkout, as `src/ui-worker.ts` does. Without presence, the UI still works and shows Offline.

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

- The grid has 16,384 columns (A to XFD) and 1,048,576 rows. A sheet name is at most 256 UTF-16 code units.
  It cannot be blank or contain `[ ] : * ? / \ ! '`. The constants `MAX_COL`, `MAX_ROW`, and `MAX_SHEET_NAME`
  hold these values.
- A cell holds at most 32,767 UTF-16 code units of text (`MAX_TEXT`). One insert or delete changes at most 10,000
  rows or columns (`MAX_STRUCTURE_COUNT`).
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
- The Worker keeps workbooks in D1 and R2. A workbook stored before sharing has no owner,
  and the Worker refuses it for every principal. `list` still names every workbook to every valid key.
  There is no rate limit yet.
- Storage uses the `Store` and `Blobs` ports of `agentic-identity`. A workbook is a document in the store
  and its state is a blob. The CLI keeps both in files under the store directory (`store.json` and
  `blobs/`). A write takes a lock file, and a lock that is more than 3 seconds old is taken over. A workbook saved by
  0.1.0 as one file per workbook moves into `store.json` on the first call, a read included. The old files stay. A 0.1.0 file whose id starts and ends with two underscores,
  or is longer than 256 characters, stops the store with an error that names the file. Rename it first.
- CSV export refuses a sheet whose extent passes 1,000,000 cells. The HTTP API refuses a write or import of
  more than 10,000 cells (413). A formula may read at most 1,000,000 range cells.

## License

MIT.
