import type { Peer, Status, SyncClient } from "agentic-realtime/client";
import * as Y from "yjs";
import type { SheetsClient } from "../src/client.js";
import type { Engine } from "../src/core/engine.js";
import type { Principal, Workbook } from "../src/core/workbook.js";

/** The presence of a room: who is here, and where their cursor is. A `SyncClient` of agentic-realtime has these. */
export type RoomPresence = Pick<SyncClient, "peers" | "onPeers" | "setPresence">;

/**
 * A workbook that the app keeps itself, in its own Yjs room. The sheet reads and writes it, and does not call a client,
 * poll, or open a presence room. The app keeps the document and its room.
 */
export interface HostWorkbook {
  /** The Yjs document. The sheet opens it as a `Workbook`: its `sheets` array, as this package stores it. */
  doc: Y.Doc;
  /** The presence of the app's room. With it, the sheet shows who is here and sends its own cursor. */
  presence?: RoomPresence;
  /** The connection state now. Without `status`, the sheet shows no connection state. */
  status?: () => Status;
  /** A way to hear when the connection state changes. It returns a function that stops the listener. The sheet reads it once each time it opens `doc`. */
  onStatus?: (listener: (status: Status) => void) => () => void;
}

/** The calls the sheet makes. A client from `createClient` has them, and so does a host workbook. */
export type SheetCalls = Pick<SheetsClient, "sheets" | "read" | "write" | "addSheet" | "deleteSheet" | "insertRows" | "deleteRows" | "insertCols" | "deleteCols">;

/**
 * A host document must be a Yjs document of the same Yjs copy as the sheet, since a document from another copy cannot be
 * read. Its presence must have peers, onPeers, and setPresence.
 */
export function checkHost(host: HostWorkbook): void {
  if (!host || !(host.doc instanceof Y.Doc)) throw new TypeError("a host workbook needs a Yjs document in `doc`");
  const presence = host.presence;
  if (presence !== undefined && (!Array.isArray(presence?.peers) || typeof presence.onPeers !== "function" || typeof presence.setPresence !== "function")) {
    throw new TypeError("a host presence needs peers, onPeers, and setPresence, as a SyncClient has");
  }
  if ((host.status !== undefined && typeof host.status !== "function") || (host.onStatus !== undefined && typeof host.onStatus !== "function")) {
    throw new TypeError("a host status and onStatus must be functions");
  }
}

/**
 * The sheet's calls, made on a workbook in this tab. Each write is stamped with `who`, so a cell an agent writes is
 * machine-made. The engine computes the values and must stay live while the calls are in use.
 */
export function workbookCalls(wb: Workbook, engine: Engine, who: Principal): SheetCalls {
  const first = (sheet?: string) => sheet ?? wb.sheetNames()[0];
  return {
    sheets: async () => ({ sheets: wb.sheetNames().map((name) => ({ name, extent: wb.extent(name) })) }),
    read: async ({ sheet, range }) => {
      const name = first(sheet);
      const cells = wb.readRange(name, range);
      return { sheet: name, range, values: engine.values(name, range), formulas: cells.map((row) => row.map((c) => c.formula)), authors: cells.map((row) => row.map((c) => c.author)) };
    },
    write: async ({ sheet, anchor, values }) => (wb.writeRange(first(sheet), anchor, values, who), { ok: true as const, rows: values.length }),
    addSheet: async ({ name }) => (wb.addSheet(name), { ok: true as const, sheets: wb.sheetNames() }),
    deleteSheet: async ({ name }) => (wb.deleteSheet(name), { ok: true as const, sheets: wb.sheetNames() }),
    insertRows: async ({ sheet, at, count = 1 }) => (wb.insertRows(first(sheet), at, count), { ok: true as const }),
    deleteRows: async ({ sheet, at, count = 1 }) => (wb.deleteRows(first(sheet), at, count), { ok: true as const }),
    insertCols: async ({ sheet, at, count = 1 }) => (wb.insertCols(first(sheet), at, count), { ok: true as const }),
    deleteCols: async ({ sheet, at, count = 1 }) => (wb.deleteCols(first(sheet), at, count), { ok: true as const }),
  } as SheetCalls;
}

export type { Peer, Status };
