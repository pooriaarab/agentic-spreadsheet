import { defineConfig } from "tsup";

export default defineConfig({
  entry: [
    "src/index.ts",
    "src/cli.ts",
    "src/worker.ts",
    "src/client.ts",
  ],
  format: ["esm"],
  dts: true,
  clean: true,
});
