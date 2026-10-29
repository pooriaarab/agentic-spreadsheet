# Changelog

All notable changes to this project are listed here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added

- `agentic-spreadsheet/client`: a typed client for the Worker HTTP API. `createClient({ baseUrl, token, fetch? })`
  has one method for each route. A failed call throws `SheetsApiError` with the HTTP `status` and a `code`.
  It uses only `fetch`, takes an `AbortSignal`, and never retries.

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
