import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["src/**/*.test.ts"],
    environment: "node",
    // Many tests spawn several CLI or MCP processes; 5 s is too tight on shared runners.
    testTimeout: 30_000,
  },
});
