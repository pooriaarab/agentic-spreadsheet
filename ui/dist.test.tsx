// @vitest-environment jsdom
// The built entry, as a host loads it. `npm run build` makes it, and ci:local builds before it tests.
import { existsSync, readFileSync } from "node:fs";
import { act } from "react";
import { describe, expect, it, vi } from "vitest";
import type { SheetsClient } from "../src/client.js";

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
window.matchMedia = ((query: string) => ({ matches: false, media: query, addEventListener() {}, removeEventListener() {} })) as unknown as typeof window.matchMedia;

const built = existsSync("dist/embed.js");
// A build is missing only on a bare `npm test`. CI builds first, and `CI` makes a missing build a failure.
if (!built && process.env.CI) throw new Error("dist/embed.js is missing: run npm run build first");

describe.skipIf(!built)("dist/embed.js and dist/embed.css", () => {
  it("renders a sheet when a host calls mount", async () => {
    const { mount } = await import("../dist/embed.js");
    const client = {
      sheets: async () => ({ sheets: [{ name: "Sheet1", extent: "A1" }] }),
      read: async () => ({ values: [[3]], formulas: [[null]], authors: [[null]] }),
    } as unknown as SheetsClient;
    const host = document.body.appendChild(document.createElement("div"));
    let sheet!: ReturnType<typeof mount>;
    await act(async () => void (sheet = mount(host, { client, docId: "b", principal: { kind: "human", name: "Ada" } })));
    await vi.waitFor(() => expect(host.querySelector('[data-pos="1,1"]')?.textContent).toBe("3"));
    act(() => sheet.unmount());
    expect(host.childElementCount).toBe(0);
  });

  it("has no CSS import, so a host without a bundler can load it", () => {
    expect(readFileSync("dist/embed.js", "utf8")).not.toMatch(/^import\s+["'][^"']*\.css["']/m);
  });

  it("ships the styles in one file that leaves the host page alone", () => {
    const css = readFileSync("dist/embed.css", "utf8");
    expect(css).not.toContain("@import");
    expect(css).toContain(".agd-button");
    expect(css).toContain("@font-face");
    // No rule that targets the whole page: bare `*`, `html`, `body`, or a bare heading, link, or paragraph.
    const selectors = [...css.matchAll(/(?:^|})\s*([^{}@]+)\{/g)].flatMap((m) => m[1].split(",").map((x) => x.trim()));
    const page = selectors.filter((x) => /^(\*|html|body|h[1-6]|p|a|code|kbd|pre|:focus-visible|::selection|\*::before|\*::after)$/.test(x));
    expect(selectors.length).toBeGreaterThan(100);
    expect(page).toEqual([]);
    // Dialogs and toasts render outside the sheet, so they carry their own color and box model.
    for (const portal of [".agd-dialog", ".agd-toast-viewport", ".agd-tooltip"]) expect(selectors).toContain(portal);
  });
});
