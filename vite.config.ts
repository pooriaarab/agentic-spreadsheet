import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

// The web UI. The Worker serves `dist/ui` (see wrangler.jsonc).
export default defineConfig({
  root: "ui",
  plugins: [react()],
  build: { outDir: "../dist/ui", emptyOutDir: true },
});
