import { defineConfig } from "vite";
import { tanstackRouter } from "@tanstack/router-plugin/vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { viteSingleFile } from "vite-plugin-singlefile";

/**
 * The app is a single page (no server of its own): `bun dev` serves it with the local token server
 * beside it. `--mode html` builds it into one standalone HTML file (dist/hub.html, routes in the URL hash).
 */
export default defineConfig(({ mode }) => ({
  server: { port: 3300 },
  // One powerbi-client: powerbi-report-authoring adds its methods (getCapabilities…) to the copy it imports.
  resolve: { tsconfigPaths: true, dedupe: ["powerbi-client"] },
  base: mode === "html" ? "./" : "/",
  build: mode === "html" ? { outDir: "dist/html", emptyOutDir: true, assetsInlineLimit: 100_000_000, cssCodeSplit: false } : undefined,
  plugins: [tanstackRouter({ target: "react", autoCodeSplitting: false }), react(), tailwindcss(), ...(mode === "html" ? [viteSingleFile()] : [])],
}));
