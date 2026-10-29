// Starts `wrangler dev` for the demo Worker. It makes .dev.vars on the first run, with a random API key.
import { spawn } from "node:child_process";
import { randomBytes } from "node:crypto";
import { existsSync, writeFileSync } from "node:fs";

if (!existsSync(".dev.vars")) {
  const key = randomBytes(24).toString("hex");
  writeFileSync(
    ".dev.vars",
    `SHEETS_API_KEYS={"${key}":{"id":"ana","kind":"human"}}\nSHEETS_REGION=unknown\nSHEETS_UNSAFE_REGION=1\n`,
  );
  console.log(`Made .dev.vars. Paste this key in the page to sign in: ${key}`);
}
const child = spawn("npx", ["wrangler", "dev", ...process.argv.slice(2)], {
  stdio: "inherit",
  env: { ...process.env, WRANGLER_SEND_METRICS: "false" },
});
child.on("exit", (code) => process.exit(code ?? 1));
