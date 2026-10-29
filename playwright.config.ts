import { defineConfig } from "@playwright/test";

// The port is picked once, and the test workers inherit it.
process.env.E2E_PORT ??= String(20000 + Math.floor(Math.random() * 20000));
const port = Number(process.env.E2E_PORT);

export default defineConfig({
  testDir: "e2e",
  testMatch: "*.spec.ts",
  timeout: 60_000,
  expect: { timeout: 10_000 },
  workers: 1,
  reporter: [["list"]],
  use: { baseURL: `http://127.0.0.1:${port}`, trace: "retain-on-failure" },
  webServer: { command: "node e2e/serve.mjs", url: `http://127.0.0.1:${port}/health`, timeout: 120_000, reuseExistingServer: false, env: { E2E_PORT: String(port) } },
});
