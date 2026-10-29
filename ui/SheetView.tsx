import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent, type KeyboardEvent } from "react";
import { BotIcon, Button, Dialog, Input, PresenceStack, useToast } from "agentic-design/react";
import { colLetters } from "../src/core/a1.js";
import type { CellAuthor, SheetsClient, Value } from "../src/client.js";
import { address, isError, parseInput, show, sizeFor, type Pos } from "./grid.js";
import { errorText, isUnauthorized } from "./session.js";
import { usePresence } from "./usePresence.js";

interface Data {
  cols: number;
  rows: number;
  values: Value[][];
  formulas: (string | null)[][];
  authors: (CellAuthor | null)[][];
}
type View = { status: "loading" } | { status: "error"; message: string } | { status: "ready"; data: Data };
// Where the text of an edit is typed: in the cell, or in the formula bar.
interface Draft {
  text: string;
  where: "cell" | "bar";
}

const POLL_MS = 4000;

export function SheetView({ client, book, token, onUnauthorized }: { client: SheetsClient; book: string; token: string; onUnauthorized: () => void }) {
  const toast = useToast();
  const [sheets, setSheets] = useState<string[]>([]);
  const [active, setActive] = useState<string | null>(null);
  const [view, setView] = useState<View>({ status: "loading" });
  const [sel, setSel] = useState<Pos>({ row: 1, col: 1 });
  const [draft, setDraft] = useState<Draft | null>(null);
  const [dialog, setDialog] = useState<"add" | "delete" | null>(null);
  const [newSheet, setNewSheet] = useState("");
  const [sheetError, setSheetError] = useState<string>();
  const grid = useRef<HTMLTableElement>(null);
  const refocus = useRef(false);
  const request = useRef(0);
  const { peers, live } = usePresence(book, token, address(sel));

  const fail = useCallback(
    (error: unknown, what: string) => {
      if (isUnauthorized(error)) return onUnauthorized();
      toast({ title: what, description: errorText(error), tone: "danger" });
    },
    [onUnauthorized, toast],
  );

  // Read the sheet list and the cells of one sheet. A silent read keeps what is on screen if it fails.
  const load = useCallback(
    async (want: string | null, silent = false) => {
      const id = ++request.current;
      if (!silent) setView({ status: "loading" });
      try {
        const listed = (await client.sheets({ book })).sheets;
        const current = listed.find((s) => s.name === want) ?? listed[0];
        const { cols, rows } = sizeFor(current.extent);
        const read = await client.read({ book, sheet: current.name, range: `A1:${colLetters(cols)}${rows}` });
        if (id !== request.current) return;
        setSheets(listed.map((s) => s.name));
        setActive(current.name);
        setView({ status: "ready", data: { cols, rows, values: read.values, formulas: read.formulas, authors: read.authors } });
      } catch (error) {
        if (isUnauthorized(error)) return onUnauthorized();
        if (!silent && id === request.current) setView({ status: "error", message: errorText(error) });
      }
    },
    [client, book, onUnauthorized],
  );
  useEffect(() => void load(null), [load]);

  const editing = draft !== null;
  useEffect(() => {
    if (editing) return;
    const timer = setInterval(() => !document.hidden && void load(active, true), POLL_MS);
    return () => clearInterval(timer);
  }, [editing, load, active]);

  useEffect(() => {
    if (!refocus.current || editing) return;
    refocus.current = false;
    grid.current?.querySelector<HTMLElement>(`[data-pos="${sel.row},${sel.col}"]`)?.focus();
  });

  const data = view.status === "ready" ? view.data : null;
  const rawAt = ({ row, col }: Pos): string => {
    if (data === null) return "";
    return data.formulas[row - 1]?.[col - 1] ?? show(data.values[row - 1]?.[col - 1] ?? null);
  };

  const move = (to: Pos) => {
    if (data === null) return;
    refocus.current = true;
    setSel({ row: Math.min(data.rows, Math.max(1, to.row)), col: Math.min(data.cols, Math.max(1, to.col)) });
  };

  // Write the draft to the selected cell, then move.
  async function commit(text: string, next: Pos) {
    setDraft(null);
    move(next);
    if (text === rawAt(sel)) return;
    try {
      await client.write({ book, sheet: active ?? undefined, anchor: address(sel), values: [[parseInput(text)]] });
      await load(active, true);
    } catch (error) {
      fail(error, `Could not save ${address(sel)}`);
    }
  }
  const cancel = () => {
    setDraft(null);
    refocus.current = true;
    setSel({ ...sel });
  };

  function onCellKey(event: KeyboardEvent<HTMLTableCellElement>) {
    if (data === null || event.target !== event.currentTarget) return;
    const { row, col } = sel;
    const keys: Record<string, Pos> = { ArrowUp: { row: row - 1, col }, ArrowDown: { row: row + 1, col }, ArrowLeft: { row, col: col - 1 }, ArrowRight: { row, col: col + 1 } };
    if (keys[event.key]) {
      event.preventDefault();
      move(keys[event.key]);
    } else if (event.key === "Enter" || event.key === "F2") {
      event.preventDefault();
      setDraft({ text: rawAt(sel), where: "cell" });
    } else if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      void commit("", sel);
    } else if (event.key === "Tab") {
      // Tab walks the cells. At the first and last cell it leaves the grid, so the grid never traps focus.
      const first = row === 1 && col === 1;
      const last = row === data.rows && col === data.cols;
      if (event.shiftKey ? first : last) return;
      event.preventDefault();
      const step = event.shiftKey ? -1 : 1;
      const wrap = col + step;
      move(wrap < 1 ? { row: row - 1, col: data.cols } : wrap > data.cols ? { row: row + 1, col: 1 } : { row, col: wrap });
    } else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      // Typing on a cell starts an edit that replaces its text.
      event.preventDefault();
      setDraft({ text: event.key, where: "cell" });
    }
  }

  function onEditKey(event: KeyboardEvent<HTMLInputElement>) {
    if (draft === null || data === null) return;
    const { row, col } = sel;
    if (event.key === "Enter") {
      event.preventDefault();
      void commit(draft.text, { row: row + 1, col });
    } else if (event.key === "Tab") {
      event.preventDefault();
      void commit(draft.text, { row, col: col + (event.shiftKey ? -1 : 1) });
    } else if (event.key === "Escape") {
      event.preventDefault();
      cancel();
    }
  }

  function onBarKey(event: KeyboardEvent<HTMLInputElement>) {
    if (event.key === "Enter" && draft !== null) {
      event.preventDefault();
      void commit(draft.text, sel);
    } else if (event.key === "Escape") {
      event.preventDefault();
      cancel();
    }
  }

  async function structure(what: string, run: () => Promise<unknown>) {
    try {
      await run();
      await load(active, true);
      refocus.current = true;
      setSel({ ...sel });
    } catch (error) {
      fail(error, what);
    }
  }
  const sheetTarget = { book, sheet: active ?? undefined };

  async function addSheet(event: FormEvent) {
    event.preventDefault();
    try {
      await client.addSheet({ book, name: newSheet.trim() });
      setDialog(null);
      setNewSheet("");
      setSheetError(undefined);
      await load(newSheet.trim());
      setSel({ row: 1, col: 1 });
    } catch (error) {
      setSheetError(errorText(error));
    }
  }
  async function deleteSheet() {
    try {
      await client.deleteSheet({ book, name: active! });
      setDialog(null);
      await load(null);
      setSel({ row: 1, col: 1 });
    } catch (error) {
      setDialog(null);
      fail(error, `Could not delete ${active}`);
    }
  }

  const cursors = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const peer of peers) if (peer.cursor !== null) map.set(peer.cursor, [...(map.get(peer.cursor) ?? []), peer.label]);
    return map;
  }, [peers]);

  const here = address(sel);
  return (
    <section className="sheet" aria-label={`Workbook ${book}`}>
      <div className="row between">
        <h1>{book}</h1>
        <div className="row">
          <span className="muted" role="status">{live === "connected" ? "Live" : live === "connecting" ? "Connecting…" : "Offline"}</span>
          <PresenceStack entries={peers.map((p) => p.entry)} label="Also here" />
        </div>
      </div>

      {view.status === "loading" && <p role="status" className="muted">Loading sheet…</p>}
      {view.status === "error" && (
        <div role="alert" className="banner">
          <p>Could not load the sheet: {view.message}</p>
          <Button size="sm" variant="secondary" onClick={() => void load(active)}>Try again</Button>
        </div>
      )}
      {data !== null && (
        <>
          <nav className="row tabs" aria-label="Sheets">
            {sheets.map((name) => (
              <Button key={name} size="sm" variant={name === active ? "primary" : "secondary"} aria-current={name === active ? "true" : undefined} onClick={() => (setSel({ row: 1, col: 1 }), void load(name))}>
                {name}
              </Button>
            ))}
            <Button size="sm" variant="ghost" onClick={() => (setSheetError(undefined), setDialog("add"))}>Add sheet</Button>
            <Button size="sm" variant="ghost" disabled={sheets.length < 2} onClick={() => setDialog("delete")}>Delete sheet</Button>
          </nav>
          <div role="group" aria-label="Rows and columns" className="row">
            <Button size="sm" variant="secondary" onClick={() => void structure("Could not insert the row", () => client.insertRows({ ...sheetTarget, at: sel.row }))}>Insert row</Button>
            <Button size="sm" variant="secondary" onClick={() => void structure("Could not delete the row", () => client.deleteRows({ ...sheetTarget, at: sel.row }))}>Delete row</Button>
            <Button size="sm" variant="secondary" onClick={() => void structure("Could not insert the column", () => client.insertCols({ ...sheetTarget, at: sel.col }))}>Insert column</Button>
            <Button size="sm" variant="secondary" onClick={() => void structure("Could not delete the column", () => client.deleteCols({ ...sheetTarget, at: sel.col }))}>Delete column</Button>
          </div>
          <div className="bar">
            <span className="namebox"><span className="sr-only">Selected cell </span>{here}</span>
            <Input
              label="Formula bar"
              autoComplete="off"
              value={draft?.text ?? rawAt(sel)}
              onFocus={() => draft === null && setDraft({ text: rawAt(sel), where: "bar" })}
              onChange={(e) => setDraft({ text: e.target.value, where: "bar" })}
              onKeyDown={onBarKey}
              onBlur={(e) => draft?.where === "bar" && e.relatedTarget !== null && void commit(draft.text, sel)}
            />
          </div>
          <div className="scroller">
            <table ref={grid} role="grid" className="grid" aria-label={`${active}, spreadsheet`} aria-rowcount={data.rows} aria-colcount={data.cols}>
              <thead>
                <tr>
                  <th scope="col" className="corner"><span className="sr-only">Row</span></th>
                  {Array.from({ length: data.cols }, (_, c) => (
                    <th key={c} scope="col" className={c + 1 === sel.col ? "on" : undefined}>{colLetters(c + 1)}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Array.from({ length: data.rows }, (_, r) => (
                  <tr key={r}>
                    <th scope="row" className={r + 1 === sel.row ? "on" : undefined}>{r + 1}</th>
                    {Array.from({ length: data.cols }, (_, c) => {
                      const pos = { row: r + 1, col: c + 1 };
                      const value = data.values[r]?.[c] ?? null;
                      const formula = data.formulas[r]?.[c] ?? null;
                      const author = data.authors[r]?.[c] ?? null;
                      const addr = address(pos);
                      const isSel = pos.row === sel.row && pos.col === sel.col;
                      const others = cursors.get(addr);
                      const label = [`${addr}: ${show(value) || "empty"}`, formula && `formula ${formula}`, author?.machine && `written by agent ${author.id}`, others && `also here: ${others.join(", ")}`].filter(Boolean).join(", ");
                      const cls = ["cell", typeof value === "number" && "num", isError(value) && "bad", author?.machine && "agent", others && "peer", isSel && "sel"].filter(Boolean).join(" ");
                      return (
                        <td key={c} role="gridcell" data-pos={`${pos.row},${pos.col}`} aria-label={editing && isSel && draft.where === "cell" ? undefined : label} aria-selected={isSel} tabIndex={isSel ? 0 : -1} className={cls} onKeyDown={onCellKey} onClick={() => !(isSel && draft?.where === "cell") && (draft !== null && void commit(draft.text, pos), draft === null && move(pos))}>
                          {isSel && draft?.where === "cell" ? (
                            <input
                              className="editor"
                              aria-label={`Edit ${addr}`}
                              autoFocus
                              value={draft.text}
                              onFocus={(e) => e.currentTarget.setSelectionRange(draft.text.length, draft.text.length)}
                              onChange={(e) => setDraft({ text: e.target.value, where: "cell" })}
                              onKeyDown={onEditKey}
                              onBlur={(e) => e.relatedTarget !== null && void commit(draft.text, sel)}
                            />
                          ) : (
                            <>
                              {show(value)}
                              {author?.machine && <BotIcon className="mark" aria-hidden="true" />}
                              {others && <span className="tag">{others[0]}</span>}
                            </>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="muted legend"><BotIcon className="mark" aria-hidden="true" /> A cell with this mark was written by an agent. Arrow keys move, Enter edits, Escape cancels, Tab moves on.</p>
        </>
      )}

      <Dialog open={dialog === "add"} onOpenChange={(open) => !open && setDialog(null)} title="Add sheet" description="Each sheet in a workbook needs its own name.">
        <form className="stack" onSubmit={addSheet}>
          <Input label="Sheet name" value={newSheet} onChange={(e) => setNewSheet(e.target.value)} error={sheetError} autoComplete="off" />
          <Button type="submit" disabled={newSheet.trim() === ""}>Add sheet</Button>
        </form>
      </Dialog>
      <Dialog
        open={dialog === "delete"}
        onOpenChange={(open) => !open && setDialog(null)}
        title={`Delete ${active}?`}
        description="A formula in another sheet that points here will read #REF!. This cannot be undone."
        footer={<Button variant="danger" onClick={() => void deleteSheet()}>Delete sheet</Button>}
      />
    </section>
  );
}

