// The spreadsheet for an app that embeds it: `agentic-spreadsheet/ui`. The app gives an element, a client, a workbook id,
// and the principal who uses the sheet. React and react-dom are peer dependencies. Add `agentic-spreadsheet/ui.css` once.
import { createRoot } from "react-dom/client";
import { ToastProvider, TooltipProvider } from "agentic-design/react";
import type { PrincipalKind, SheetsClient } from "../src/client.js";
import { SheetView } from "./SheetView.js";

export type { PrincipalKind, SheetsClient };

const PRINCIPAL_KINDS: readonly PrincipalKind[] = ["human", "agent_of_human", "agent_of_agent", "agent", "user_token"];

/** Who uses the sheet. The server still checks every write against its own token. */
export interface EmbedPrincipal {
  /** Checked only. It does not change what the server allows. */
  kind: PrincipalKind;
  name: string;
}

export interface MountOptions {
  /** Reads and writes the workbook. A client from `createClient` in `agentic-spreadsheet/client` works. */
  client: SheetsClient;
  /** The workbook id. */
  docId: string;
  principal: EmbedPrincipal;
  /** Show the cells and allow no edit. The default is false. */
  readOnly?: boolean;
  /** The key for the live presence room, on the origin of the page. Without it, the sheet shows no peers. */
  token?: string;
  /** Called when the server does not accept the client's key. */
  onUnauthorized?: () => void;
}

/** A principal must have a known kind and a name. A missing kind is an error, never a default. */
function checkOptions(options: MountOptions): void {
  const { client, docId } = options;
  if (typeof client?.sheets !== "function" || typeof client.read !== "function") throw new TypeError("the sheet needs a client from createClient");
  if (typeof docId !== "string" || docId === "") throw new TypeError("the sheet needs a docId that is a non-empty string");
  checkPrincipal(options.principal);
}

function checkPrincipal(principal: EmbedPrincipal): void {
  if (!principal || !PRINCIPAL_KINDS.includes(principal.kind) || typeof principal.name !== "string" || principal.name.trim() === "") {
    throw new TypeError(`the sheet needs a principal with a kind (${PRINCIPAL_KINDS.join(", ")}) and a name`);
  }
}

const noop = () => {};

/** The grid of one workbook, with its tabs, formula bar, and presence. It needs the design providers above it. */
export function WorkbookView({ client, docId, principal, readOnly = false, token, onUnauthorized = noop }: MountOptions) {
  checkOptions({ client, docId, principal });
  return <SheetView client={client} book={docId} token={token} name={principal.name} readOnly={readOnly} onUnauthorized={onUnauthorized} />;
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
