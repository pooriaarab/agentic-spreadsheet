# agentic-spreadsheet rules

These rules bind every agent that edits a workbook with this package.

1. Edit as yourself. Pass your own principal. Do not pass `human` for agent work.
2. Read before you write. Check the range you plan to overwrite.
3. Report failures as failures. A thrown error stays an error.
