import { test, type Page } from "@playwright/test";
import { createClient } from "../src/client.js";
import { ANA } from "./keys.mjs";

// Screenshots for a design review. They run only when SHOTS_DIR names a folder: SHOTS_DIR=/tmp/shots npx playwright test shots
const dir = process.env.SHOTS_DIR;
const origin = `http://127.0.0.1:${process.env.E2E_PORT}`;

async function open(page: Page, book: string, theme: "light" | "dark") {
  await page.setViewportSize({ width: 1200, height: 700 });
  await page.emulateMedia({ colorScheme: theme, reducedMotion: "reduce" });
  await page.goto("/");
  if (await page.getByLabel("API key").isVisible()) {
    await page.getByLabel("API key").fill(ANA);
    await page.getByRole("button", { name: "Sign in" }).click();
  }
  await page.getByLabel("Theme").selectOption(theme);
  await page.goto(`/#/${book}`);
  await page.locator('td[data-pos="1,1"]').waitFor();
}

test("screenshots", async ({ page }) => {
  test.skip(!dir, "set SHOTS_DIR");
  const book = `shots-${Date.now().toString(36)}`;
  const api = createClient({ baseUrl: origin, token: ANA });
  await api.create({ book });
  await api.write({ book, anchor: "A1", values: [["Item", "Qty", "Price", "Total"], ["Pens", 12, 1.5, "=B2*C2"], ["Paper", 4, 6, "=B3*C3"], [null, null, "Sum", "=SUM(D2:D3)"]] });
  await open(page, book, "light");
  await page.locator('td[data-pos="2,4"]').click();
  await page.keyboard.press("ArrowLeft");
  await page.screenshot({ path: `${dir}/grid-light.png` });
  await open(page, book, "dark");
  await page.locator('td[data-pos="2,4"]').click();
  await page.screenshot({ path: `${dir}/grid-dark.png` });
});
