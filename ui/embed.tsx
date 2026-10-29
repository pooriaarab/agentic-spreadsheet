// The spreadsheet for an app that embeds it: `agentic-spreadsheet/ui`. The app gives an element, a client, a workbook id,
// and the principal who uses the sheet. Or it gives its own Yjs document as `host`. React and react-dom are peer
// dependencies. Add `agentic-spreadsheet/ui.css` once.
import { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { ToastProvider, TooltipProvider } from "agentic-design/react";
import type { PrincipalKind, SheetsClient } from "../src/client.js";
import { Engine } from "../src/core/engine.js";
import { Workbook } from "../src/core/workbook.js";
import { checkHost, workbookCalls, type HostWorkbook, type RoomPresence, type SheetCalls } from "./host.js";
import { SheetView } from "./SheetView.js";

export type { HostWorkbook, PrincipalKind, RoomPresence, SheetCalls, SheetsClient };

const PRINCIPAL_KINDS: readonly PrincipalKind[] = ["human", "agent_of_human", "agent_of_agent", "agent", "user_token"];

/** Who uses the sheet. The server still checks every write against its own token. */
export interface EmbedPrincipal {
  /** Checked only. It does not change what the server allows. */
  kind: PrincipalKind;
  name: string;
  /** The id stamped on each cell this principal writes in a host workbook. The default is `name`. */
  id?: string;
}

export interface MountOptions {
  /** Reads and writes the workbook. A client from `createClient` in `agentic-spreadsheet/client` works. Not needed with `host`. */
  client?: SheetsClient;
  /** The workbook id. */
  docId: string;
  principal: EmbedPrincipal;
  /** Show the cells and allow no edit. The default is false. */
  readOnly?: boolean;
  /** The key for the live presence room, on the origin of the page. Without it, the sheet shows no peers. Not allowed with `host`. */
  token?: string;
  /**
   * A Yjs document that the app keeps in its own room. The sheet reads and writes it, and does not call a client, poll,
   * or open a presence room. The app keeps the document and its room.
   */
  host?: HostWorkbook;
  /** The heading of the sheet. The default is the docId. */
  title?: string;
  /** Called when the server does not accept the client's key. */
  onUnauthorized?: () => void;
}

/** A principal must have a known kind and a name. A missing kind is an error, never a default. */
function checkOptions(options: MountOptions): void {
  const { client, docId, host } = options;
  if (host !== undefined) {
    checkHost(host);
    if (options.token !== undefined) throw new TypeError("the sheet takes a token or a host workbook, not both");
  } else if (typeof client?.sheets !== "function" || typeof client.read !== "function") {
    throw new TypeError("the sheet needs a client from createClient, or a host workbook");
  }
  if (typeof docId !== "string" || docId === "") throw new TypeError("the sheet needs a docId that is a non-empty string");
  checkPrincipal(options.principal);
}

function checkPrincipal(principal: EmbedPrincipal): void {
  if (!principal || !PRINCIPAL_KINDS.includes(principal.kind) || typeof principal.name !== "string" || principal.name.trim() === "") {
    throw new TypeError(`the sheet needs a principal with a kind (${PRINCIPAL_KINDS.join(", ")}) and a name`);
  }
  if (principal.id !== undefined && (typeof principal.id !== "string" || principal.id.trim() === "")) throw new TypeError("a principal id must be a non-empty string");
}

const noop = () => {};

/** The grid of one workbook, with its tabs, formula bar, and presence. It needs the design providers above it. */
export function WorkbookView(options: MountOptions) {
  checkOptions(options);
  const { client, docId, principal, readOnly = false, token, host, title, onUnauthorized = noop } = options;
  if (host) return <HostedSheet host={host} docId={docId} principal={principal} readOnly={readOnly} title={title} />;
  return <SheetView client={client!} book={docId} title={title} token={token} name={principal.name} readOnly={readOnly} onUnauthorized={onUnauthorized} />;
}

/** The sheet on a workbook the app keeps. The workbook and its engine live while the document does not change. */
function HostedSheet({ host, docId, principal, readOnly, title }: { host: HostWorkbook; docId: string; principal: EmbedPrincipal; readOnly: boolean; title?: string }) {
  const [open, setOpen] = useState<{ book: Workbook; engine: Engine } | null>(null);
  // Made and disposed in one effect, so a second mount (StrictMode) gets a live pair, and none stays on the host's document.
  useEffect(() => {
    const book = new Workbook(host.doc);
    const engine = new Engine(book);
    setOpen({ book, engine });
    return () => {
      engine.dispose();
      book.dispose();
    };
  }, [host.doc]);
  const { kind, name, id = name } = principal;
  const calls = useMemo(() => open && workbookCalls(open.book, open.engine, { id, kind }), [open, id, kind]);
  const doc = host.doc;
  const changes = useMemo(() => (listener: () => void) => (doc.on("update", listener), () => doc.off("update", listener)), [doc]);
  // The app may make the host object on each render. The connection is read once each time the document changes.
  const latest = useRef(host);
  latest.current = host;
  const hasStatus = host.status !== undefined;
  const connection = useMemo(
    () => (hasStatus ? { status: () => latest.current.status!(), onStatus: (listener: Parameters<NonNullable<HostWorkbook["onStatus"]>>[0]) => latest.current.onStatus?.(listener) ?? noop } : undefined),
    [doc, hasStatus],
  );
  if (!calls || open?.book.doc !== doc) return <p role="status" className="as-muted">Loading sheet…</p>;
  return <SheetView key={docId} client={calls} book={docId} title={title} changes={changes} presence={host.presence} connection={connection} name={name} readOnly={readOnly} onUnauthorized={noop} />;
}

export interface MountedSheet {
  /** Show another workbook, principal, or mode in the same place. A bad option throws and keeps the sheet as it is. After `unmount`, it does nothing. */
  update(options: MountOptions): void;
  /** Remove the sheet from the element and stop its polling and presence. Safe to call twice. */
  unmount(): void;
}

/** Render the sheet into an element, for an app that does not use React itself. */
export function mount(element: HTMLElement, options: MountOptions): MountedSheet {
  if (!(element instanceof HTMLElement)) throw new TypeError("mount needs an element to render into");
  checkOptions(options);
  const root = createRoot(element);
  let open = true;
  const render = (next: MountOptions) =>
    root.render(
      <TooltipProvider>
        <ToastProvider>
          <WorkbookView key={next.docId} {...next} />
        </ToastProvider>
      </TooltipProvider>,
    );
  render(options);
  return {
    update: (next) => (checkOptions(next), open && render(next)),
    unmount: () => {
      if (!open) return;
      open = false;
      root.unmount();
    },
  };
}
