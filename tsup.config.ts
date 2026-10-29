import { build } from "esbuild";
import { defineConfig } from "tsup";

export default defineConfig({
  // Named entries, so `ui/embed.tsx` lands next to the others as `dist/embed.js`.
  entry: {
    index: "src/index.ts",
    cli: "src/cli.ts",
    worker: "src/worker.ts",
    client: "src/client.ts",
    embed: "ui/embed.tsx",
  },
  format: ["esm"],
  // The styles of `./ui`, with the tokens and the self-hosted fonts of agentic-design, in one file: dist/embed.css.
  // The JS has no CSS import, so any host can load it, with or without a bundler.
  onSuccess: async () => {
    await build({ entryPoints: ["ui/app.css"], bundle: true, loader: { ".woff2": "file" }, outfile: "dist/embed.css", logLevel: "warning" });
  },
  // The UI entry is JSX and types HTMLElement. The root tsconfig is for Node code.
  dts: { compilerOptions: { jsx: "react-jsx", lib: ["ES2022", "DOM", "DOM.Iterable"] } },
  clean: true,
});
