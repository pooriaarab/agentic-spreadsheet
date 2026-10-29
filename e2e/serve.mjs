// Start `wrangler dev` for the E2E tests: the API, the UI files in dist/ui, and the presence room.
// Data lives in memory, so each run starts empty. No telemetry, and a private home for wrangler.
import { spawn } from "node:child_process";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { KEYS } from "./keys.mjs";

const port = Number(process.env.E2E_PORT);
const wrangler = fileURLToPath(new URL("../node_modules/wrangler/bin/wrangler.js", import.meta.url));
const home = mkdtempSync(join(tmpdir(), "wrangler-ui-"));
const child = spawn(
  process.execPath,
  [wrangler, "dev", "--port", String(port), "--inspector-port", String(port + 1), "--ip", "127.0.0.1", "--var", `SHEETS_API_KEYS:${JSON.stringify(KEYS)}`],
  { stdio: "inherit", env: { ...process.env, CI: "1", NO_COLOR: "1", WRANGLER_SEND_METRICS: "false", XDG_CONFIG_HOME: home, HOME: home } },
);
for (const signal of ["SIGINT", "SIGTERM"]) process.on(signal, () => child.kill(signal));
child.on("exit", (code) => process.exit(code ?? 0));
