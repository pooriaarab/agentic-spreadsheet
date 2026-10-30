import { readFileSync } from "node:fs";
import { describe, expect, test } from "vitest";
import { VERSION } from "./version.js";

describe("VERSION", () => {
  test("matches package.json without a build", () => {
    const pkg = JSON.parse(
      readFileSync(new URL("../package.json", import.meta.url), "utf8"),
    ) as { version: string };
    expect(VERSION).toBe(pkg.version);
  });

  test("matches the MCP registry manifest", () => {
    const server = JSON.parse(
      readFileSync(new URL("../server.json", import.meta.url), "utf8"),
    ) as { version: string; packages: { version: string }[] };
    expect(server.version).toBe(VERSION);
    expect(server.packages.map((pkg) => pkg.version)).toEqual([VERSION]);
  });
});
