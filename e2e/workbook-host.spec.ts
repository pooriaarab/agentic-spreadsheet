import AxeBuilder from "@axe-core/playwright";
import { expect, test, type Browser, type Page } from "@playwright/test";
import { build } from "esbuild";
import { createClient } from "../src/client.js";
import { ANA, BEN, BOT } from "./keys.mjs";

// The file name sorts after ui.spec.ts on purpose: its first test needs a server with no workbooks.

// Ways the host option can fail, listed before the tests:
// - an edit in one host does not reach the other host through the host's room, or a formula does not compute
// - a peer in the host's room does not show, or an agent's cell looks like a person's
// - the sheet polls or calls the API when a host gives its document
// - unmount closes the host's room or drops the host's document
// - the sheet on a host page has an axe violation of WCAG 2.2 AA in light or dark

const origin = `http://127.0.0.1:${process.env.E2E_PORT}`;
const api = (key: string) => createClient({ baseUrl: origin, token: key });
const unique = (name: string) => `${name}-${Date.now().toString(36)}${Math.floor(Math.random() * 1000)}`;
const TAGS = ["wcag2a", "wcag2aa", "wcag21a", "wcag21aa", "wcag22aa"];
const pos = (addr: string) => {
  const [, col, row] = /^([A-Z]+)(\d+)$/.exec(addr)!;
  return `${row},${[...col].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)}`;
};
const cell = (page: Page, addr: string) => page.locator(`td[data-pos="${pos(addr)}"]`);

// The host page is built from e2e/host/main.ts with the built `agentic-spreadsheet/ui`, and served on the Worker's origin.
// The embed styles leave the page alone, so the host page sets its own background from the design tokens.
let script = "";
let styles = "";
test.beforeAll(async () => {
  const out = await build({ entryPoints: ["e2e/host/main.ts"], bundle: true, format: "esm", target: "es2022", write: false, outdir: "out", define: { "process.env.NODE_ENV": '"production"' }, logLevel: "warning" });
  script = out.outputFiles.find((f) => f.path.endsWith(".js"))!.text;
  styles = (await build({ entryPoints: ["dist/embed.css"], bundle: true, write: false, loader: { ".woff2": "dataurl" }, logLevel: "warning" })).outputFiles[0].text;
});

async function host(browser: Browser, key: string, book: string, who: Record<string, string>, theme: "light" | "dark" = "light"): Promise<Page> {
  const page = await (await browser.newContext()).newPage();
  await page.route("**/e2e-host/**", (route) =>
    new URL(route.request().url()).pathname.endsWith(".js") ? route.fulfill({ contentType: "text/javascript", body: script }) : route.fulfill({ contentType: "text/css", body: styles }),
  );
  await page.emulateMedia({ colorScheme: theme, reducedMotion: "reduce" });
  // A page from the Worker itself, so the browser counts the host page as local and lets it open the room's socket.
  await page.goto(`/health#${new URLSearchParams({ key, book, ...who })}`);
  await page.setContent(
    `<!doctype html><html lang="en" data-theme="${theme}"><head><meta charset="utf-8"><title>Host workbook</title><link rel="stylesheet" href="/e2e-host/app.css"><style>body { margin: 1rem; background: var(--ad-color-bg); color: var(--ad-color-text); }</style></head><body><main><div id="sheet"></div></main><script type="module" src="/e2e-host/app.js"></script></body></html>`,
  );
  await expect(page.getByRole("grid")).toBeVisible();
  return page;
}

async function typeIn(page: Page, addr: string, text: string) {
  await cell(page, addr).click();
  await page.keyboard.type(text);
  await page.keyboard.press("Enter");
}

test("two host apps share one workbook through their own room, see each other, and keep the document after unmount", async ({ browser }) => {
  const book = unique("host");
  await api(ANA).create({ book });
  await api(ANA).share({ book, kind: "human", subject: "ben", role: "editor" });
  const ana = await host(browser, ANA, book, { kind: "human", name: "Ana", id: "ana", seed: "1" });
  const ben = await host(browser, BEN, book, { kind: "human", name: "Ben", id: "ben" });
  // The sheet makes no API call with a host document: no poll, no read, no write.
  const calls: string[] = [];
  for (const page of [ana, ben]) page.on("request", (request) => request.url().includes("/api/v1/sheets") && calls.push(request.url()));

  await expect(ana.getByRole("heading", { name: "Budget", level: 1 })).toBeVisible();
  await expect(ana.getByText("Live")).toBeVisible();
  await typeIn(ana, "A1", "21");
  await typeIn(ben, "B1", "=A1*2");
  await expect(cell(ana, "B1")).toHaveText("42");
  await expect(cell(ben, "A1")).toHaveText("21");
  await expect(ana.getByRole("list", { name: "Also here" }).getByRole("listitem")).toHaveCount(1);
  await cell(ben, "C3").click();
  await expect(cell(ana, "C3")).toHaveAttribute("aria-label", /also here/);
  await new Promise((r) => setTimeout(r, 4500)); // longer than the poll of a sheet on a client
  expect(calls).toEqual([]);

  // The sheet goes. The host's room stays connected, and its document still gets Ana's edits.
  await ben.evaluate(() => (window as unknown as { unmountSheet: () => void }).unmountSheet());
  await expect(ben.getByRole("grid")).toHaveCount(0);
  await expect(cell(ana, "C3")).not.toHaveAttribute("aria-label", /also here/);
  expect(await ben.evaluate(() => (window as unknown as { hostStatus: () => string }).hostStatus())).toBe("connected");
  await typeIn(ana, "A2", "7");
  await expect.poll(() => ben.evaluate(() => (window as unknown as { hostCell: (a: string) => unknown }).hostCell("A2"))).toBe(7);
  for (const page of [ana, ben]) await page.context().close();
});

for (const theme of ["light", "dark"] as const) {
  test(`on a host workbook an agent's cell is marked, a read-only sheet writes nothing, and axe passes in ${theme}`, async ({ browser }) => {
    const book = unique(`host-${theme}`);
    await api(ANA).create({ book });
    await api(ANA).share({ book, kind: "agent", subject: "bot-1", role: "editor" });
    const ana = await host(browser, ANA, book, { kind: "human", name: "Ana", id: "ana", seed: "1" }, theme);
    const bot = await host(browser, BOT, book, { kind: "agent_of_human", name: "Scribe", id: "bot-1" }, theme);
    await expect(ana.locator(".ad-stack__summary")).toContainText("1 agent");
    await typeIn(bot, "A1", "5");
    await expect(cell(ana, "A1")).toHaveAttribute("aria-label", /written by agent bot-1/);
    await expect(cell(ana, "A1").locator(".as-mark")).toHaveCount(1);
    await cell(ana, "A1").click();
    const { violations } = await new AxeBuilder({ page: ana }).withTags(TAGS).analyze();
    expect(violations.map((v) => `${v.id}: ${v.nodes.map((n) => n.target.join(" ")).join(" | ")}`)).toEqual([]);

    const viewer = await host(browser, ANA, book, { kind: "human", name: "Ana", id: "ana", readonly: "1" }, theme);
    await cell(viewer, "B1").click();
    await viewer.keyboard.type("9");
    await viewer.keyboard.press("Enter");
    await expect(viewer.locator("input.as-editor")).toHaveCount(0);
    await expect(cell(ana, "B1")).toHaveAttribute("aria-label", /B1: empty/);
    for (const page of [ana, bot, viewer]) await page.context().close();
  });
}
