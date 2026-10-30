import { readFileSync } from "node:fs";
import ts from "typescript";
import { describe, expect, it } from "vitest";
import pkg from "../package.json" with { type: "json" };

// Match the declaration entry names in package.json to their source files.
const sources: Record<string, string> = {
  "./dist/index.d.ts": "src/index.ts",
  "./dist/client.d.ts": "src/client.ts",
  "./dist/formats.d.ts": "src/formats.ts",
  "./dist/embed.d.ts": "ui/embed.tsx",
  "./dist/worker.d.ts": "src/worker.ts",
  "./dist/account.d.ts": "src/account.ts",
};

describe("public export documentation", () => {
  it("names each export under its package path in the README", () => {
    const readme = readFileSync("README.md", "utf8");
    const entries = Object.entries(pkg.exports);
    const program = ts.createProgram(Object.values(sources), {
      module: ts.ModuleKind.NodeNext,
      moduleResolution: ts.ModuleResolutionKind.NodeNext,
      jsx: ts.JsxEmit.ReactJSX,
      target: ts.ScriptTarget.ES2022,
      skipLibCheck: true,
    });
    const checker = program.getTypeChecker();
    for (const [path, target] of entries) {
      const name = path === "." ? pkg.name : pkg.name + path.slice(1);
      const heading = `### \`${name}\``;
      const section = readme.split(heading)[1]?.split("\n### ")[0]?.split("\n## ")[0];
      expect(section, `Missing section for ${name}`).toBeDefined();
      if (typeof target === "string") continue; // CSS has no named exports.
      const source = sources[target.types];
      expect(source, `Map the declaration entry for ${name} to its source`).toBeDefined();
      const file = program.getSourceFile(source!);
      expect(file, `Missing source for ${name}`).toBeDefined();
      const module = checker.getSymbolAtLocation(file!);
      expect(module, `Missing module for ${name}`).toBeDefined();
      for (const symbol of checker.getExportsOfModule(module!)) {
        expect(section, `${name} is missing ${symbol.name}`).toContain(`\`${symbol.name}\``);
      }
    }
  });
});
