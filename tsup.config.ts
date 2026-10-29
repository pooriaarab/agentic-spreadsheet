import { readFile } from "node:fs/promises";
import { dirname } from "node:path";
import { build, type Plugin } from "esbuild";
import { defineConfig } from "tsup";

// The base rules of agentic-design style `*`, `body`, headings, and links. In a host page they would restyle the host,
// so the embed build drops them. `ui/app.css` scopes the few it needs under `.as-sheet`.
const withoutPageRules: Plugin = {
  name: "without-page-rules",
  setup: (b) =>
    b.onLoad({ filter: /agentic-design[\\/]dist[\\/]styles\.css$/ }, async ({ path }) => {
      const css = await readFile(path, "utf8");
      const start = css.indexOf("/* Base rules.");
      const end = css.indexOf("/* Component styles.");
      if (start < 0 || end < start) throw new Error("agentic-design/styles.css changed: update withoutPageRules in tsup.config.ts");
      const keep = css.slice(start, end).match(/\.ad-sr-only \{[^}]*\}/)?.[0] ?? "";
      return { contents: css.slice(0, start) + keep + "\n" + css.slice(end), loader: "css", resolveDir: dirname(path) };
    }),
};

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
  esbuildOptions: (options) => {
    options.jsx = "automatic";
  },
  // The styles of `./ui`, with the tokens and the self-hosted fonts of agentic-design, in one file: dist/embed.css.
  // The JS has no CSS import, so any host can load it, with or without a bundler.
  onSuccess: async () => {
    await build({ entryPoints: ["ui/app.css"], bundle: true, loader: { ".woff2": "file" }, outfile: "dist/embed.css", logLevel: "warning", plugins: [withoutPageRules] });
  },
  // The UI entry is JSX and types HTMLElement. The root tsconfig is for Node code.
  dts: { compilerOptions: { jsx: "react-jsx", lib: ["ES2022", "DOM", "DOM.Iterable"] } },
  clean: true,
});
