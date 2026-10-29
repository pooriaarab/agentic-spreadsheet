// A host app that keeps the workbook in its own Yjs room, and gives it to `agentic-spreadsheet/ui` as `host`.
import { SyncClient, type Status } from "agentic-realtime/client";
import { mount } from "agentic-spreadsheet/ui";
import { Workbook } from "../../src/core/workbook.js";
import * as Y from "yjs";

const params = new URLSearchParams(location.hash.slice(1));
const key = params.get("key") ?? "";
const book = params.get("book") ?? "";
const principal = { kind: params.get("kind") as "human", name: params.get("name") ?? "", id: params.get("id") ?? undefined };
const base64url = (text: string) => btoa(text).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
// The host's own socket. The room reads the key from a subprotocol.
class HostSocket extends WebSocket {
  constructor(url: string | URL) {
    super(url, ["agentic-spreadsheet", `agentic-spreadsheet.bearer.${base64url(key)}`]);
  }
}

const url = new URL(`/api/v1/docs/${encodeURIComponent(book)}/ws`, location.href);
url.protocol = url.protocol === "https:" ? "wss:" : "ws:";
const doc = new Y.Doc();
let status: Status = "connecting";
const listeners = new Set<(status: Status) => void>();
const sync = SyncClient.open({ url: url.toString(), WebSocket: HostSocket as never, doc, presence: { name: principal.name }, reconnect: {},
  onStatus: (next) => {
    status = next;
    for (const listener of listeners) listener(next);
  },
});
await sync.whenSynced();
// The host makes the first sheet, as an app does when it makes the file.
const seed = new Workbook(doc);
if (params.has("seed") && seed.sheetNames().length === 0) seed.addSheet("Sheet1");
seed.dispose();
const sheet = mount(document.getElementById("sheet")!, {
  docId: book,
  title: "Budget",
  principal,
  readOnly: params.has("readonly"),
  host: { doc, presence: sync, status: () => status, onStatus: (listener) => (listeners.add(listener), () => void listeners.delete(listener)) },
});
Object.assign(window, {
  unmountSheet: () => sheet.unmount(),
  // The host still holds the document and the room after the sheet goes.
  hostCell: (address: string) => {
    const read = new Workbook(doc);
    try {
      return read.readRange("Sheet1", address)[0][0].value;
    } finally {
      read.dispose();
    }
  },
  hostStatus: () => status,
});
