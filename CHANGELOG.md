# Changelog

All notable changes to this project are listed here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

## [0.5.0] - 2026-10-29

### Breaking changes

- A delete of rows or columns that runs past the last row or column of the grid (row 1,048,576 or column XFD)
  is refused with an error. Before, the call gave no error. Check the range before you delete.
- A sheet name longer than 256 characters is refused when a sheet is added or renamed. Before, any length
  was accepted. Names that already exist are not changed.
- Deploying the Worker from a checkout needs `npm run build` before `wrangler dev` or `wrangler deploy`,
  because `wrangler.jsonc` serves `dist/ui`. The first deploy also applies a new migration for the
  `SYNC_ROOM` Durable Object. A path outside `/api/` and `/health` now gets the UI page, not 404.
  The default export of `agentic-spreadsheet/worker` does not change.

### Added

- A web UI, served by the Worker: workbook list, grid with a formula bar, edits, structure changes, presence,
  and an agent marker. It runs on `agentic-design`, the typed client, and `agentic-realtime`. See the README, "UI".
  The presence socket is `/api/v1/docs/:book/ws`. The browser sends the key in its `token` query parameter, so a
  key can appear in request URLs and logs. Presence runs only in the Worker in `src/ui-worker.ts`, which the npm
  package does not ship.
- Operations `delete-rows`, `insert-cols`, `delete-cols`, `sheets`, `add-sheet`, and `delete-sheet`. Each has a
  CLI command, an MCP tool (`sheet_` prefix), an HTTP route, and a client method.
  - HTTP: `DELETE /api/v1/sheets/:book/rows?at=&count=`, `POST` and `DELETE /api/v1/sheets/:book/cols?at=&count=`,
    `GET` and `POST /api/v1/sheets/:book/sheets`, and `DELETE /api/v1/sheets/:book/sheets/:name`.
- Client methods `deleteRows`, `insertCols`, `deleteCols`, `sheets`, `addSheet`, and `deleteSheet`.
- A named export `roomGrant` from `agentic-spreadsheet/worker`. It checks who may join a workbook's presence room.
- The SDK exports `MAX_SHEET_NAME` (256), the longest sheet name that is accepted.

### Changed

- `npm run build` also builds the UI into `dist/ui`. `npm run typecheck` also checks the UI.
- `ci:local` also runs the Playwright end-to-end tests against `wrangler dev`, with an axe WCAG 2.2 AA
  check in light and dark. A machine that runs them needs the Playwright browser.
- New scripts: `npm run dev:ui` starts `wrangler dev`, and `npm run test:e2e` runs the Playwright tests.
- The package has no new runtime dependencies. `agentic-design`, `agentic-realtime`, `react`, and
  `react-dom` are development dependencies. They are built into `dist/ui`.

## [0.4.0] - 2026-10-29

### Added

- `agentic-spreadsheet/client`: a typed client for the Worker HTTP API. `createClient({ baseUrl, token, fetch? })`
  has one method for each route. A failed call throws `SheetsApiError` with the HTTP `status` and a `code`.
  It uses only `fetch`, takes an `AbortSignal`, and never retries.

### Breaking changes

- A workbook with a cell key in the old A1 form on a sheet that has an id now refuses reads and writes, with an
  error that says how to recover. Before, the workbook loaded and the cell read as empty. Restore the workbook
  from a copy that has no such key.
- An old formula that holds U+0001 or U+0002 is refused when it is converted, because the stored form reads
  those characters as ref markers.

### Changed

- A sheet that comes without an id after a conversion converts with a tag of its own name. It is a clash only
  when two copies convert that same sheet.
- `DELETE /api/v1/sheets/:book` deletes the blob, then erases the grants, then deletes the document. A failed
  step can run again with the owner's grant in place, and the name stays taken until the end.
- A store conflict during that delete answers 409. A malformed blob field in a stored record is a data error,
  and the Worker answers 500.

### Fixed

- A failed blob delete in `DELETE /api/v1/sheets/:book` is no longer hidden. The caller gets a server error, with
  no backend text, and the same call can run again.
- A save that lands during a delete makes the call start again. The grants step runs once.

### Tests

- Every read-only command is tested to save nothing, by comparing the stored bytes and the file list.

## [0.3.0] - 2026-10-29

### Breaking changes

- `agentic-identity` is now 0.4.0. A role belongs to a kind and an id. A person and an agent with the same id
  hold separate roles. A role that a 0.2.0 Worker gave to an id does not carry over.
  This includes the owner role. A workbook made by a 0.2.0 Worker in a persistent `Store` has no owner after
  the upgrade, and there is no migration.
- `PUT /api/v1/sheets/:book/share` needs a `kind` (`human` or `agent`) in the body. The take-back route is now
  `DELETE /api/v1/sheets/:book/share/:kind/:subject`. A principal of any kind but `human` is an agent.

### Added

- `DELETE /api/v1/sheets/:book` lets the owner delete a workbook. It also erases the grants, invites, links,
  and audit entries of the workbook with `deleteResource`, so no grant outlives it.

## [0.2.0] - 2026-10-29

### Breaking changes

- Node 22 or later is required. Version 0.1.0 said Node 18. `agentic-identity` 0.3.0 needs Node 22.
- The CLI file store has a new layout. A workbook was one file per workbook. Now the store keeps
  `store.json` and a `blobs/` folder. A 0.1.0 workbook is read as it is and moves into `store.json`
  on the first write. A 0.1.0 workbook whose id starts and ends with two underscores, or is longer
  than 256 characters, stops the store with an error that names the file. Rename the file, then try again.
- The Worker checks a role on every workbook route. In 0.1.0 any valid key could use any workbook.
  Now a key with no share gets 404, even for a workbook another key made. A role that is too low gets 403.
  Only the owner shares. A workbook stored before sharing has no owner, and the Worker refuses it for every principal.
- `createApp` and `Ctx` take a `blobs` port. Pass `createMemoryBlobs()` or `createFileBlobs(dir)`, as well as the `Store`.
- `list` sorts ids by code unit, and a cursor is opaque. Pass back the `cursor` that `list` returned.
- A new workbook id may not start and end with two underscores, and may not pass 256 characters.
- A writer that meets other writers gives up after 50 tries with 409 "Try again".
- Removed exports: `MemoryStore`, `FileStore`, `queryDocs`, `SAFE_NAME`, and the types `Doc`,
  `QueryOptions`, and `QueryResult`. Use `createMemoryStore` and `createFileStore`.
- A custom `Store` must match the `agentic-identity` port. It has `region` in place of `jurisdiction`.
  A query returns `nextCursor`, not `cursor`. A transaction may run its function again, or throw `StoreConflictError`.

### Added

- Each new workbook records its creator as its owner. The owner shares it with a principal id and a role:
  `viewer`, `commenter`, or `editor`. The owner takes a share back.
- `PUT /api/v1/sheets/:book/share` gives a role. `DELETE /api/v1/sheets/:book/share/:subject` takes it back.

### Changed

- Storage uses the `Store` and `Blobs` ports of `agentic-identity`. The CLI file store keeps a
  document file and a blob folder. A workbook saved by 0.1.0 is read as it is.
- `list` sorts ids by code unit (`B` before `a`), as the `Store` port does. With no limit it returns every id.
- The file adapter writes under the store directory only on a write, and a broken 0.1.0 workbook file is an error.
- A writer that keeps meeting other writers gives up after 50 tries and answers 409 "Try again". Version 0.1.0 waited on one lock.
- A new workbook id may not start and end with two underscores, and may not pass 256 characters. These are the port rules.
- A workbook list cursor is opaque. Pass back the `cursor` that `list` returned.

### Removed

- `MemoryStore`, `FileStore`, `queryDocs`, `SAFE_NAME`, and the types `Doc`, `QueryOptions`, and `QueryResult`. See Breaking changes.

## [0.1.0] - 2026-10-29

First release.

### Added

- Workbook model on one Yjs document. Sheets have a set order. Cells are sparse.
- Every row and column has a stable id. Cells and formula refs use the ids, and A1 is a view.
- Insert, delete, and move rows and columns. Formula refs follow the moved cells.
- Formula engine with incremental recalculation and the error values `#REF!`, `#DIV/0!`,
  `#NAME?`, `#VALUE!`, `#NUM!`, `#N/A`, `#CIRCULAR`, and `#DEPTH!`.
- CSV import and export. Export marks text that could run as a formula.
- Every cell keeps the principal that last edited it.
- CLI (`agentic-spreadsheet`) and MCP server (`agentic-spreadsheet mcp`) over a local file store.
- HTTP API in a Cloudflare Worker under `/api/v1/sheets`. It needs an API key or the explicit
  `SHEETS_UNSAFE_TRUST_HEADERS=1` setting.
- Workbooks saved by an older version with A1 keys convert when they load.

### Known limits

- The Worker keeps workbooks in memory. A restart loses them.
- XLSX and ODS import and export are not built yet.
- The CLI and the MCP server do not check the caller. `--as` and `--kind` are labels.

See the README for the full list of limits.
