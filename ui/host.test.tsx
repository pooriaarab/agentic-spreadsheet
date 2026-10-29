// @vitest-environment jsdom
import type { Peer, Status } from "agentic-realtime/client";
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import * as Y from "yjs";
import type { SheetsClient } from "../src/client.js";
import { Workbook } from "../src/core/workbook.js";
import { mount, type HostWorkbook, type MountedSheet, type MountOptions } from "./embed.js";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
// jsdom has no layout, and the grid scrolls its cells into view.
Element.prototype.scrollIntoView = () => {};
window.matchMedia = ((query: string) => ({ matches: false, media: query, addEventListener() {}, removeEventListener() {} })) as unknown as typeof window.matchMedia;

const ada = { kind: "human" as const, name: "Ada", id: "ada" };

/** A host document with one sheet: A1 is 7, B1 is =A1*2, written by a person. */
function hostDoc(fill = true) {
  const doc = new Y.Doc();
  const book = new Workbook(doc);
  if (fill) {
    book.addSheet("Sheet1");
    book.writeRange("Sheet1", "A1", [[7, "=A1*2"]], { id: "seed", kind: "human" });
  }
  return { doc, book };
}

/** A client that fails the test when the sheet calls it. With a host document the sheet never calls one. */
const untouchable = () =>
  new Proxy({} as SheetsClient, {
    get: (_t, name) => (name === "then" ? undefined : vi.fn(async () => Promise.reject(new Error(`the sheet called client.${String(name)}`)))),
  });

function fakePresence() {
  const listeners = new Set<(peers: Peer[]) => void>();
  const presence = {
    peers: [] as Peer[],
    onPeers: vi.fn((listener: (peers: Peer[]) => void) => (listeners.add(listener), () => void listeners.delete(listener))),
    setPresence: vi.fn(),
    disconnect: vi.fn(),
  };
  const setPeers = (peers: Peer[]) => {
    presence.peers = peers;
    for (const listener of listeners) listener(peers);
  };
  return { presence, setPeers, listeners };
}

let mounted: MountedSheet | undefined;
let element: HTMLElement;
afterEach(() => {
  act(() => mounted?.unmount());
  mounted = undefined;
  element?.remove();
});

async function show(options: Omit<MountOptions, "principal" | "docId"> & { principal?: MountOptions["principal"]; docId?: string }) {
  element = document.body.appendChild(document.createElement("div"));
  await act(async () => void (mounted = mount(element, { principal: ada, docId: "budget", ...options })));
}
const cell = (pos: string) => element.querySelector<HTMLElement>(`[data-pos="${pos}"]`)!;
const button = (name: string) => [...element.querySelectorAll("button")].find((b) => b.textContent === name)!;
const ready = () => vi.waitFor(() => expect(element.querySelector('[role="grid"]')).not.toBeNull());

async function typeInto(pos: string, text: string) {
  const target = cell(pos);
  await act(async () => target.click());
  await act(async () => target.focus());
  await act(async () => void target.dispatchEvent(new KeyboardEvent("keydown", { key: text[0], bubbles: true })));
  const editor = element.querySelector<HTMLInputElement>("input.as-editor")!;
  await act(async () => {
    const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!;
    set.call(editor, text);
    editor.dispatchEvent(new Event("input", { bubbles: true }));
  });
  await act(async () => void editor.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })));
}

describe("mount with a host document", () => {
  it("shows the host's cells with computed formulas, and never calls a client", async () => {
    const { doc } = hostDoc();
    const client = untouchable();
    await show({ client, host: { doc } });
    await ready();
    expect(cell("1,1").textContent).toBe("7");
    expect(cell("1,2").textContent).toBe("14");
    expect(Object.values(client).length).toBe(0);
  });

  it("needs no client", async () => {
    await show({ host: { doc: hostDoc().doc } });
    await ready();
    expect(cell("1,1").textContent).toBe("7");
  });

  it("writes an edit into the host's document, stamped with the principal", async () => {
    const { doc, book } = hostDoc();
    await show({ host: { doc } });
    await ready();
    await typeInto("1,1", "9");
    expect(book.readRange("Sheet1", "A1")[0][0]).toMatchObject({ value: 9, author: { id: "ada", kind: "human", machine: false } });
    await vi.waitFor(() => expect(cell("1,2").textContent).toBe("18"));
  });

  it("an agent principal's cells are marked as machine-made", async () => {
    const { doc, book } = hostDoc();
    await show({ host: { doc }, principal: { kind: "agent_of_human", name: "Scribe", id: "bot-1" } });
    await ready();
    await typeInto("2,1", "5");
    expect(book.readRange("Sheet1", "A2")[0][0]).toMatchObject({ value: 5, author: { id: "bot-1", kind: "agent_of_human", machine: true } });
    await vi.waitFor(() => expect(cell("2,1").getAttribute("aria-label")).toContain("written by agent bot-1"));
  });

  it("an open edit is not moved to another sheet when a peer renames the sheet it is on", async () => {
    const { doc, book } = hostDoc();
    book.addSheet("Q2");
    const peer = new Y.Doc();
    Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
    peer.on("update", (update: Uint8Array) => Y.applyUpdate(doc, update));
    await show({ host: { doc } });
    await ready();
    await act(async () => button("Q2").click());
    await vi.waitFor(() => expect(element.querySelector('[aria-current="true"]')?.textContent).toBe("Q2"));
    await act(async () => cell("1,1").focus());
    await act(async () => void cell("1,1").dispatchEvent(new KeyboardEvent("keydown", { key: "5", bubbles: true })));
    const other = new Workbook(peer);
    await act(async () => other.renameSheet("Q2", "Q2 final"));
    other.dispose();
    const editor = element.querySelector<HTMLInputElement>("input.as-editor")!;
    await act(async () => void editor.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })));
    expect(book.readRange("Sheet1", "A1")[0][0].value).toBe(7);
    expect(book.readRange("Q2 final", "A1")[0][0].value).toBe(null);
    await vi.waitFor(() => expect(document.body.textContent).toContain("Could not save A1"));
  });

  it("a workbook that cannot open shows an error in the sheet, not a broken page", async () => {
    const doc = new Y.Doc();
    doc.getArray("sheets").push(["not a sheet"] as never);
    await show({ host: { doc } });
    await vi.waitFor(() => expect(element.querySelector('[role="alert"]')?.textContent).toContain("Could not open the sheet"));
  });

  it("redraws at once when the host's document changes elsewhere, with no polling", async () => {
    const { doc } = hostDoc();
    const peer = new Y.Doc();
    Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
    peer.on("update", (update: Uint8Array) => Y.applyUpdate(doc, update));
    await show({ host: { doc } });
    await ready();
    vi.useFakeTimers({ toFake: ["setInterval"] }); // a poll would need the timers to run
    try {
      await act(async () => new Workbook(peer).writeRange("Sheet1", "A1", [[100]], { id: "ben", kind: "human" }));
      await vi.waitFor(() => expect(cell("1,2").textContent).toBe("200"));
    } finally {
      vi.useRealTimers();
    }
  });

  it("inserts rows and adds sheets in the host's document", async () => {
    const { doc, book } = hostDoc();
    await show({ host: { doc } });
    await ready();
    await act(async () => cell("1,1").click());
    await act(async () => button("Insert row").click());
    await vi.waitFor(() => expect(book.readRange("Sheet1", "A2")[0][0].value).toBe(7));
    await act(async () => button("Add sheet").click());
    const input = document.querySelector<HTMLInputElement>('input[autocomplete="off"]:not([aria-label="Formula bar"])')!;
    const name = [...document.querySelectorAll("label")].find((l) => l.textContent === "Sheet name")!.control as HTMLInputElement;
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!.call(name ?? input, "Notes");
      (name ?? input).dispatchEvent(new Event("input", { bubbles: true }));
    });
    await act(async () => (name ?? input).form!.requestSubmit());
    await vi.waitFor(() => expect(book.sheetNames()).toEqual(["Sheet1", "Notes"]));
  });

  it("a read-only sheet on a host document writes nothing", async () => {
    const { doc } = hostDoc();
    const updates = vi.fn();
    await show({ host: { doc }, readOnly: true });
    await ready();
    doc.on("update", updates);
    await act(async () => void cell("1,1").dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true })));
    expect(element.querySelector("input.as-editor")).toBeNull();
    expect(updates).not.toHaveBeenCalled();
  });

  it("an empty host workbook shows a status, and the grid once the host adds a sheet", async () => {
    const { doc, book } = hostDoc(false);
    await show({ host: { doc } });
    await vi.waitFor(() => expect(element.textContent).toContain("This workbook has no sheet yet."));
    expect(element.querySelector('[role="alert"]')).toBeNull();
    await act(async () => void book.addSheet("First"));
    await ready();
  });

  it("shows the host's peers and their cells, sends its cursor, and on unmount clears it and keeps the document", async () => {
    const { doc } = hostDoc();
    const { presence, setPeers, listeners } = fakePresence();
    const destroyed = vi.fn();
    doc.on("destroy", destroyed);
    // Yjs keeps the deep observers of a type here. The workbook and the engine must leave none behind.
    const observers = () => (doc.getArray("sheets") as unknown as { _dEH: { l: unknown[] } })._dEH.l.length;
    const before = observers();
    await show({ host: { doc, presence } });
    await ready();
    expect(presence.setPresence).toHaveBeenLastCalledWith({ cursor: "A1" });
    await act(async () => setPeers([{ clientId: 3, name: "Scribe", machine: true, kind: "agent", cursor: "B1", color: "", away: "active" }]));
    expect(cell("1,2").getAttribute("aria-label")).toContain("also here");
    expect(element.querySelector(".agd-stack__summary")?.textContent).toBe("1 agent");
    await act(async () => cell("2,1").click());
    expect(presence.setPresence).toHaveBeenLastCalledWith({ cursor: "A2" });
    act(() => mounted!.unmount());
    mounted = undefined;
    expect(presence.setPresence).toHaveBeenLastCalledWith({ cursor: null });
    expect(listeners.size).toBe(0);
    expect(presence.disconnect).not.toHaveBeenCalled();
    expect(destroyed).not.toHaveBeenCalled();
    expect(observers()).toBe(before);
  });

  it("a workbook that fails to convert leaves no observer on the host's document", async () => {
    // An old workbook (a sheet with no id) whose edit log forks. The conversion refuses it.
    const doc = new Y.Doc();
    const old = new Y.Map<unknown>();
    old.set("name", "Sheet1");
    doc.getArray("sheets").push([old]);
    doc.getArray("structure").push([{ id: "a", parent: null }, { id: "b", parent: "x" }]);
    const observers = () => (doc.getArray("sheets") as unknown as { _dEH: { l: unknown[] } })._dEH.l.length;
    await show({ host: { doc } });
    await vi.waitFor(() => expect(element.querySelector('[role="alert"]')?.textContent).toContain("Could not open the sheet"));
    expect(observers()).toBe(0);
  });

  it("a read-only sheet does not convert an old workbook, and says why", async () => {
    const doc = new Y.Doc();
    const old = new Y.Map<unknown>();
    old.set("name", "Sheet1");
    old.set("cells", new Y.Map());
    doc.getArray("sheets").push([old]);
    const updates = vi.fn();
    doc.on("update", updates);
    await show({ host: { doc }, readOnly: true });
    await vi.waitFor(() => expect(element.querySelector('[role="alert"]')?.textContent).toContain("old format"));
    expect(updates).not.toHaveBeenCalled();
  });

  it("peers go when the host's presence goes", async () => {
    const { doc } = hostDoc();
    const { presence, setPeers } = fakePresence();
    await show({ host: { doc, presence } });
    await ready();
    await act(async () => setPeers([{ clientId: 3, name: "Scribe", machine: true, kind: "agent", cursor: "B1", color: "", away: "active" }]));
    expect(element.querySelector(".agd-stack__summary")?.textContent).toBe("1 agent");
    await act(async () => mounted!.update({ docId: "budget", principal: ada, host: { doc } }));
    expect(element.querySelector(".agd-stack__summary")?.textContent ?? "").not.toContain("agent");
    expect(cell("1,2").getAttribute("aria-label")).not.toContain("also here");
  });

  it("shows the host's connection state", async () => {
    const listeners = new Set<(status: Status) => void>();
    let status: Status = "connected";
    await show({ host: { doc: hostDoc().doc, status: () => status, onStatus: (l) => (listeners.add(l), () => void listeners.delete(l)) } });
    await ready();
    expect(element.textContent).toContain("Live");
    await act(async () => {
      status = "disconnected";
      for (const l of listeners) l(status);
    });
    expect(element.textContent).toContain("Offline");
    act(() => mounted!.unmount());
    mounted = undefined;
    expect(listeners.size).toBe(0);
  });

  it("uses the title for the heading when one is given", async () => {
    await show({ host: { doc: hostDoc().doc }, title: "Budget 2026" });
    await ready();
    expect(element.querySelector("h1")?.textContent).toBe("Budget 2026");
  });

  it("update moves the sheet to another host document", async () => {
    await show({ host: { doc: hostDoc().doc } });
    await ready();
    const other = hostDoc(false);
    other.book.addSheet("Other");
    other.book.writeRange("Other", "A1", [["moved"]], { id: "x", kind: "human" });
    await act(async () => mounted!.update({ docId: "budget", principal: ada, host: { doc: other.doc } }));
    await vi.waitFor(() => expect(cell("1,1").textContent).toBe("moved"));
  });

  it("refuses a host with a token, a host principal with no id, a host with no Yjs document, a bad presence, and no host with no client", () => {
    element = document.createElement("div");
    const { doc } = hostDoc();
    const bad: object[] = [{ host: { doc }, token: "key" }, { host: { doc }, principal: { kind: "human", name: "Ada" } }, { host: {} }, { host: { doc: { getArray: () => null } } }, { host: { doc, presence: { peers: [] } } }, {}];
    for (const options of bad) expect(() => mount(element, { docId: "b", principal: ada, ...options } as never)).toThrow(TypeError);
    expect(element.childElementCount).toBe(0);
  });

  it("update with a bad host throws and keeps the sheet", async () => {
    await show({ host: { doc: hostDoc().doc } });
    await ready();
    expect(() => mounted!.update({ docId: "budget", principal: ada, host: {} as HostWorkbook })).toThrow(TypeError);
    expect(cell("1,1").textContent).toBe("7");
  });
});
