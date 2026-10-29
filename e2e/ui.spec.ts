import AxeBuilder from "@axe-core/playwright";
import { expect, test, type Page } from "@playwright/test";
import { createClient } from "../src/client.js";
import { ANA, BEN, BOT } from "./keys.mjs";

// Ways the UI can fail, listed before the tests:
// - a wrong key signs in, or the page shows a blank screen instead of an error
// - a value or formula the user types does not reach the API, or the computed result does not show, or a reload loses it
// - a key does something the grid should not: Escape saves, Enter does not edit, Tab traps focus
// - a row, column, or sheet change does not reach the API, or the grid shows the old cells
// - a cell an agent wrote looks like a person's cell, or only color tells them apart
// - a peer's presence never shows, or shows after the peer left
// - a load that fails or hangs shows nothing, or a failed request shows as success
// - a view has an axe violation of WCAG 2.2 AA in light or dark

const origin = `http://127.0.0.1:${process.env.E2E_PORT}`;
const api = (key: string) => createClient({ baseUrl: origin, token: key });
const unique = (name: string) => `${name}-${Date.now().toString(36)}${Math.floor(Math.random() * 1000)}`;
const TAGS = ["wcag2a", "wcag2aa", "wcag21a", "wcag21aa", "wcag22aa"];
const THEMES = ["light", "dark"] as const;

async function signIn(page: Page, key = ANA) {
  await page.goto("/");
  await page.getByLabel("API key").fill(key);
  await page.getByRole("button", { name: "Sign in" }).click();
  await expect(page.getByRole("button", { name: "New workbook" }).first()).toBeVisible();
}
async function createBook(page: Page, name: string) {
  await page.getByRole("button", { name: "New workbook" }).first().click();
  await page.getByLabel("Workbook name").fill(name);
  await page.getByRole("button", { name: "Create workbook" }).click();
  await expect(page.getByRole("heading", { name, level: 1 })).toBeVisible();
  await expect(cell(page, "A1")).toBeVisible();
}
const cell = (page: Page, addr: string) => page.locator(`td[data-pos="${pos(addr)}"]`);
const pos = (addr: string) => {
  const [, col, row] = /^([A-Z]+)(\d+)$/.exec(addr)!;
  return `${row},${[...col].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)}`;
};
async function audit(page: Page, what: string) {
  const result = await new AxeBuilder({ page }).withTags(TAGS).analyze();
  expect(result.violations.map((v) => `${what}: ${v.id} ${v.help} (${v.nodes.map((n) => n.target.join(" ")).join("; ")})`)).toEqual([]);
}
async function inTheme(page: Page, theme: (typeof THEMES)[number]) {
  await page.emulateMedia({ colorScheme: theme, reducedMotion: "reduce" });
  await page.getByLabel("Theme").selectOption(theme);
  await expect(page.locator("html")).toHaveAttribute("data-theme", theme);
}

test("a wrong key is refused with a message, and a right key opens the app", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Sign in to Sheets" })).toBeVisible();
  await page.getByLabel("API key").fill("x".repeat(40));
  await page.getByRole("button", { name: "Sign in" }).click();
  await expect(page.getByRole("alert")).toContainText("not valid");
  await signIn(page);
  await expect(page.getByText("No workbooks yet")).toBeVisible();
});

test("create a sheet, enter a value and a formula, see the result, and reload keeps it", async ({ page }) => {
  const book = unique("budget");
  await signIn(page);
  await createBook(page, book);
  await cell(page, "A1").click();
  await page.keyboard.type("5");
  await page.keyboard.press("Enter");
  await expect(cell(page, "A1")).toHaveText("5");
  await expect(cell(page, "A2")).toBeFocused();
  await page.keyboard.press("ArrowUp");
  await page.keyboard.press("ArrowRight");
  await page.keyboard.type("=A1*2");
  await page.keyboard.press("Tab");
  await expect(cell(page, "B1")).toHaveText("10");
  await expect(cell(page, "C1")).toBeFocused();
  // The API holds the same cells.
  const read = await api(ANA).read({ book, range: "A1:B1" });
  expect(read.values).toEqual([[5, 10]]);
  expect(read.formulas).toEqual([[null, "=A1*2"]]);

  await page.reload();
  await expect(cell(page, "B1")).toHaveText("10");
  await cell(page, "B1").click();
  await expect(page.getByLabel("Formula bar")).toHaveValue("=A1*2");
});

test("keys: Enter edits, Escape cancels, typing replaces, Delete clears, arrows move, focus shows", async ({ page }) => {
  const book = unique("keys");
  await api(ANA).create({ book });
  await api(ANA).write({ book, anchor: "A1", values: [[1, 2]] });
  await signIn(page);
  await page.goto(`/#/${book}`);
  await expect(cell(page, "A1")).toHaveText("1");
  await cell(page, "A1").focus();
  await page.keyboard.press("Enter");
  await page.keyboard.type("99");
  await page.keyboard.press("Escape");
  await expect(cell(page, "A1")).toHaveText("1");
  await expect(cell(page, "A1")).toBeFocused();
  // Focus is visible: the focused cell has an outline.
  expect(await cell(page, "A1").evaluate((el) => getComputedStyle(el).outlineStyle)).toBe("solid");
  await page.keyboard.press("ArrowRight");
  await expect(cell(page, "B1")).toBeFocused();
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowLeft");
  await expect(cell(page, "A2")).toBeFocused();
  await page.keyboard.press("ArrowUp");
  await page.keyboard.press("Enter");
  await page.keyboard.press("Control+a");
  await page.keyboard.type("7");
  await page.keyboard.press("Enter");
  await expect(cell(page, "A1")).toHaveText("7");
  await page.keyboard.press("ArrowUp");
  await page.keyboard.press("Delete");
  await expect(cell(page, "A1")).toHaveText("");
  expect((await api(ANA).read({ book, range: "A1:B1" })).values).toEqual([[null, 2]]);
  // Tab walks to the next cell, and Shift+Tab walks back.
  await page.keyboard.press("Tab");
  await expect(cell(page, "B1")).toBeFocused();
  await page.keyboard.press("Shift+Tab");
  await expect(cell(page, "A1")).toBeFocused();
  // The formula bar edits the selected cell.
  await page.getByLabel("Formula bar").fill("=B1+40");
  await page.getByLabel("Formula bar").press("Enter");
  await expect(cell(page, "A1")).toHaveText("42");
});

test("add and remove rows, columns, and sheets", async ({ page }) => {
  const book = unique("shape");
  await api(ANA).create({ book });
  await api(ANA).write({ book, anchor: "A1", values: [[1, 2], [3, "=A2*2"]] });
  await signIn(page);
  await page.goto(`/#/${book}`);
  await expect(cell(page, "B2")).toHaveText("6");
  await cell(page, "A1").click();
  await page.getByRole("button", { name: "Insert row" }).click();
  await expect(cell(page, "A2")).toHaveText("1");
  await expect(cell(page, "B3")).toHaveText("6");
  await page.getByRole("button", { name: "Delete row" }).click();
  await expect(cell(page, "A1")).toHaveText("1");
  await page.getByRole("button", { name: "Insert column" }).click();
  await expect(cell(page, "B1")).toHaveText("1");
  await page.getByRole("button", { name: "Delete column" }).click();
  await expect(cell(page, "A1")).toHaveText("1");
  expect((await api(ANA).read({ book, range: "A1:B2" })).values).toEqual([[1, 2], [3, 6]]);

  await page.getByRole("button", { name: "Add sheet" }).click();
  await page.getByLabel("Sheet name").fill("Data");
  await page.getByRole("button", { name: "Add sheet" }).last().click();
  await expect(page.getByRole("button", { name: "Data" })).toHaveAttribute("aria-current", "true");
  expect((await api(ANA).sheets({ book })).sheets.map((s) => s.name)).toEqual(["Sheet1", "Data"]);
  await page.getByRole("button", { name: "Delete sheet" }).click();
  await page.getByRole("dialog").getByRole("button", { name: "Delete sheet" }).click();
  await expect(page.getByRole("button", { name: "Data" })).toHaveCount(0);
  expect((await api(ANA).sheets({ book })).sheets.map((s) => s.name)).toEqual(["Sheet1"]);
});

test("a cell an agent wrote carries a mark and a name that is not color alone", async ({ page }) => {
  const book = unique("agent");
  await api(ANA).create({ book });
  await api(ANA).share({ book, kind: "agent", subject: "bot-1", role: "editor" });
  await api(BOT).write({ book, anchor: "A1", values: [[3]] });
  await api(ANA).write({ book, anchor: "B1", values: [[4]] });
  await signIn(page);
  await page.goto(`/#/${book}`);
  await expect(cell(page, "A1")).toHaveAttribute("aria-label", /written by agent bot-1/);
  await expect(cell(page, "A1").locator("svg")).toHaveCount(1);
  await expect(cell(page, "B1")).not.toHaveAttribute("aria-label", /agent/);
  await expect(cell(page, "B1").locator("svg")).toHaveCount(0);
  // An agent edit that lands while the page is open shows up with no reload.
  await api(BOT).write({ book, anchor: "C1", values: [["=A1+B1"]] });
  await expect(cell(page, "C1")).toHaveText("7", { timeout: 15_000 });
  await expect(cell(page, "C1")).toHaveAttribute("aria-label", /written by agent bot-1/);
});

test("two people in one workbook see each other", async ({ browser }) => {
  const book = unique("room");
  await api(ANA).create({ book });
  await api(ANA).share({ book, kind: "human", subject: "ben", role: "editor" });
  const [one, two] = await Promise.all([browser.newContext(), browser.newContext()]);
  const [a, b] = [await one.newPage(), await two.newPage()];
  await signIn(a, ANA);
  await signIn(b, BEN);
  for (const page of [a, b]) await page.goto(`/#/${book}`);
  await expect(a.getByText("Live")).toBeVisible();
  await expect(b.getByText("Live")).toBeVisible();
  await expect(a.getByRole("list", { name: "Also here" })).toBeVisible();
  await cell(b, "C3").click();
  await expect(cell(a, "C3")).toHaveAttribute("aria-label", /also here/);
  await two.close();
  await expect(a.getByRole("list", { name: "Also here" }).getByRole("listitem")).toHaveCount(0, { timeout: 15_000 });
  await one.close();
});

test("a failed load shows an error with a retry, and a slow load shows a status", async ({ page }) => {
  await signIn(page);
  let fail = true;
  await page.route("**/api/v1/sheets", async (route) => {
    if (route.request().method() !== "GET") return route.continue();
    if (fail) return route.fulfill({ status: 500, json: { error: "something failed on the server" } });
    await new Promise((r) => setTimeout(r, 800));
    return route.continue();
  });
  await page.reload();
  await expect(page.getByRole("alert")).toContainText("something failed on the server");
  await audit(page, "error view");
  fail = false;
  await page.getByRole("button", { name: "Try again" }).click();
  await expect(page.getByText("Loading workbooks…")).toBeVisible();
  await expect(page.getByRole("alert")).toHaveCount(0);
});

test("a write the API refuses shows an error and does not look saved", async ({ page }) => {
  const book = unique("deny");
  await api(ANA).create({ book });
  await api(ANA).share({ book, kind: "human", subject: "ben", role: "viewer" });
  await signIn(page, BEN);
  await page.goto(`/#/${book}`);
  await cell(page, "A1").click();
  await page.keyboard.type("9");
  await page.keyboard.press("Enter");
  // Exact: the status region also holds this text, and a substring match finds two elements.
  await expect(page.getByText("the viewer role cannot write this workbook", { exact: true })).toBeVisible();
  await expect(cell(page, "A1")).toHaveText("");
});

for (const theme of THEMES) {
  test(`every view passes axe WCAG 2.2 AA in ${theme}, and the screenshots are saved`, async ({ page }) => {
    const shot = async (name: string) => {
      if (process.env.UPDATE_SCREENSHOTS === "1") await page.screenshot({ path: `docs/screenshots/${name}-${theme}.png`, animations: "disabled" });
    };
    await page.emulateMedia({ colorScheme: theme, reducedMotion: "reduce" });
    await page.goto("/");
    await expect(page.getByRole("heading", { name: "Sign in to Sheets" })).toBeVisible();
    await audit(page, `sign-in ${theme}`);
    await shot("1-sign-in");

    await signIn(page);
    await inTheme(page, theme);
    // The server holds the books of earlier tests, so show the empty view with an empty list.
    await page.route("**/api/v1/sheets", (route) => (route.request().method() === "GET" ? route.fulfill({ json: { books: [] } }) : route.continue()));
    await page.reload();
    await expect(page.getByText("No workbooks yet")).toBeVisible();
    await audit(page, `empty ${theme}`);
    await shot("2-empty");
    await page.unroute("**/api/v1/sheets");
    await page.reload();
    const book = unique(`view-${theme}`);
    await page.getByRole("button", { name: "New workbook" }).first().click();
    await expect(page.getByRole("dialog")).toBeVisible();
    await audit(page, `new workbook dialog ${theme}`);
    await page.getByLabel("Workbook name").fill(book);
    await page.getByRole("button", { name: "Create workbook" }).click();
    await expect(cell(page, "A1")).toBeVisible();
    await api(ANA).share({ book, kind: "agent", subject: "bot-1", role: "editor" });
    await api(BOT).write({ book, anchor: "A1", values: [["Item", "Cost", "Note"], ["Rent", 1200, "=B2*12"], ["Food", 400, "=B3*12"], ["Total", "=B2+B3", "=B4*12"]] });
    await api(ANA).write({ book, anchor: "A6", values: [["Bad", "=1/0"]] });
    await expect(cell(page, "C4")).toHaveText("19200", { timeout: 15_000 });
    await cell(page, "B2").click();
    await audit(page, `grid ${theme}`);
    await shot("3-grid");
    await page.getByRole("button", { name: "Add sheet" }).click();
    await audit(page, `add sheet dialog ${theme}`);
    await page.keyboard.press("Escape");
  });
}
