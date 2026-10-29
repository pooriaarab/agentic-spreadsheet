# Changelog

All notable changes to this project are listed here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Changed

- Storage uses the `Store` and `Blobs` ports of `agentic-identity`. The CLI file store keeps a
  document file and a blob folder. A workbook saved by 0.1.0 is read as it is.
- `list` sorts ids by code unit (`B` before `a`), as the `Store` port does. With no limit it returns every id.
- The file adapter writes under the store directory only on a write, and a broken 0.1.0 workbook file is an error.
- A writer that keeps meeting other writers gives up after 50 tries and answers 409 "Try again". Version 0.1.0 waited on one lock.
- A workbook list cursor is opaque. Pass back the `cursor` that `list` returned.

### Removed

- `MemoryStore`, `FileStore`, `queryDocs`, `SAFE_NAME`, and the types `Doc`, `QueryOptions`, and `QueryResult`. Use `createMemoryStore`, `createFileStore`, and the `Store` type.

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
