// @vitest-environment jsdom
import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { SheetsApiError, type SheetsClient } from "../src/client.js";
import { mount, type MountedSheet, type MountOptions } from "./embed.js";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
// jsdom has no layout, and the grid scrolls its cells into view.
Element.prototype.scrollIntoView = () => {};
window.matchMedia = ((query: string) => ({ matches: false, media: query, addEventListener() {}, removeEventListener() {} })) as unknown as typeof window.matchMedia;

const ada = { kind: "human" as const, name: "Ada" };

/** A client with one sheet. A1 is 7 and B1 is a formula. */
function fakeClient(overrides: Partial<Record<keyof SheetsClient, unknown>> = {}) {
  const calls = { write: vi.fn(async () => ({ ok: true })), read: vi.fn(), sheets: vi.fn() };
  const client = {
    sheets: calls.sheets.mockImplementation(async () => ({ sheets: [{ name: "Sheet1", extent: "B1" }] })),
    read: calls.read.mockImplementation(async () => ({
      values: [[7, 14]],
      formulas: [[null, "=A1*2"]],
      authors: [[null, { id: "bot", machine: true }]],
    })),
    write: calls.write,
    insertRows: vi.fn(),
    ...overrides,
  } as unknown as SheetsClient;
  return { client, ...calls };
}

let mounted: MountedSheet | undefined;
let host: HTMLElement;
afterEach(() => {
  act(() => mounted?.unmount());
  mounted = undefined;
  host?.remove();
  vi.useRealTimers();
});

async function show(options: Omit<MountOptions, "principal"> & { principal?: MountOptions["principal"] }) {
  host = document.body.appendChild(document.createElement("div"));
  await act(async () => void (mounted = mount(host, { principal: ada, ...options })));
  await vi.waitFor(() => expect(host.querySelector('[role="grid"]')).not.toBeNull());
}
const cell = (pos: string) => host.querySelector<HTMLElement>(`[data-pos="${pos}"]`)!;
const button = (name: string) => [...host.querySelectorAll("button")].find((b) => b.textContent === name)!;

describe("mount", () => {
  it("renders the cells of the workbook into the element", async () => {
    const { client, read } = fakeClient();
    await show({ client, docId: "budget" });
    expect(host.querySelector("h1")?.textContent).toBe("budget");
    expect(cell("1,1").textContent).toBe("7");
    expect(cell("1,2").textContent).toBe("14");
    expect(cell("1,2").getAttribute("aria-label")).toContain("written by agent bot");
    expect(read).toHaveBeenCalledWith(expect.objectContaining({ book: "budget", sheet: "Sheet1" }));
  });

  it("writes an edit through the client", async () => {
    const { client, write } = fakeClient();
    await show({ client, docId: "budget" });
    const a1 = cell("1,1");
    await act(async () => a1.focus());
    await act(async () => void a1.dispatchEvent(new KeyboardEvent("keydown", { key: "9", bubbles: true })));
    const editor = host.querySelector<HTMLInputElement>("input.as-editor")!;
    expect(editor.value).toBe("9");
    await act(async () => void editor.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })));
    expect(write).toHaveBeenCalledWith({ book: "budget", sheet: "Sheet1", anchor: "A1", values: [[9]] });
  });

  it("opens read only: no editor, no write, disabled structure buttons", async () => {
    const { client, write } = fakeClient();
    await show({ client, docId: "budget", readOnly: true });
    const a1 = cell("1,1");
    for (const key of ["Enter", "F2", "Delete", "x"]) await act(async () => void a1.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true })));
    expect(host.querySelector("input.as-editor")).toBeNull();
    expect(write).not.toHaveBeenCalled();
    expect(host.querySelector('[role="grid"]')?.getAttribute("aria-readonly")).toBe("true");
    expect(host.querySelector<HTMLInputElement>('input[aria-label="Formula bar"]')?.readOnly).toBe(true);
    for (const name of ["Insert row", "Delete row", "Insert column", "Delete column", "Add sheet"]) expect(button(name).disabled).toBe(true);
  });

  it("refuses a principal with no kind, a bad kind, or no name", () => {
    host = document.createElement("div");
    const { client } = fakeClient();
    const bad = [{ name: "Ada" }, { kind: "root", name: "Ada" }, { kind: "agent", name: " " }, undefined];
    for (const principal of bad) expect(() => mount(host, { client, docId: "b", principal: principal as never })).toThrow(TypeError);
    expect(host.childElementCount).toBe(0);
  });

  it("refuses a bad element, client, or docId", () => {
    const { client } = fakeClient();
    expect(() => mount(null as never, { client, docId: "b", principal: ada })).toThrow(TypeError);
    host = document.createElement("div");
    for (const docId of ["", 5 as never]) expect(() => mount(host, { client, docId, principal: ada })).toThrow(TypeError);
    expect(() => mount(host, { client: {} as never, docId: "b", principal: ada })).toThrow(TypeError);
  });

  it("closes an open edit when the sheet turns read only, and keeps the sheet on another callback", async () => {
    const { client, write, sheets } = fakeClient();
    await show({ client, docId: "budget" });
    const a1 = cell("1,1");
    await act(async () => void a1.dispatchEvent(new KeyboardEvent("keydown", { key: "9", bubbles: true })));
    expect(host.querySelector("input.as-editor")).not.toBeNull();
    const loads = sheets.mock.calls.length;
    await act(async () => mounted!.update({ client, docId: "budget", principal: ada, readOnly: true, onUnauthorized: () => {} }));
    expect(host.querySelector("input.as-editor")).toBeNull();
    expect(write).not.toHaveBeenCalled();
    expect(sheets.mock.calls.length).toBe(loads);
  });

  it("shows the error and offers a retry when the workbook cannot load", async () => {
    const { client, sheets } = fakeClient();
    sheets.mockRejectedValueOnce(new Error("boom"));
    host = document.body.appendChild(document.createElement("div"));
    await act(async () => void (mounted = mount(host, { client, docId: "b", principal: ada })));
    await vi.waitFor(() => expect(host.querySelector('[role="alert"]')?.textContent).toContain("boom"));
    await act(async () => button("Try again").click());
    await vi.waitFor(() => expect(host.querySelector('[role="grid"]')).not.toBeNull());
  });

  it("calls onUnauthorized when the server refuses the key", async () => {
    const denied = new SheetsApiError(401, "no");
    const { client } = fakeClient({ sheets: vi.fn(async () => Promise.reject(denied)) });
    const onUnauthorized = vi.fn();
    host = document.body.appendChild(document.createElement("div"));
    await act(async () => void (mounted = mount(host, { client, docId: "b", principal: ada, onUnauthorized })));
    await vi.waitFor(() => expect(onUnauthorized).toHaveBeenCalled());
  });

  it("update shows another workbook; a bad principal keeps the sheet", async () => {
    const { client, read } = fakeClient();
    await show({ client, docId: "one" });
    expect(() => mounted!.update({ client, docId: "two", principal: { name: "x" } as never })).toThrow(TypeError);
    expect(host.querySelector("h1")?.textContent).toBe("one");
    await act(async () => mounted!.update({ client, docId: "two", principal: ada }));
    await vi.waitFor(() => expect(host.querySelector("h1")?.textContent).toBe("two"));
    expect(read).toHaveBeenLastCalledWith(expect.objectContaining({ book: "two" }));
  });

  it("unmount empties the element, stops polling, and can run twice", async () => {
    vi.useFakeTimers({ toFake: ["setInterval", "clearInterval"] });
    const { client, sheets } = fakeClient();
    await show({ client, docId: "budget" });
    const before = sheets.mock.calls.length;
    act(() => mounted!.unmount());
    expect(host.childElementCount).toBe(0);
    expect(() => mounted!.unmount()).not.toThrow();
    await vi.advanceTimersByTimeAsync(20_000);
    expect(sheets.mock.calls.length).toBe(before);
    mounted!.update({ client, docId: "x", principal: ada });
    expect(host.childElementCount).toBe(0);
  });
});
