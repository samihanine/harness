import { defineConfig } from "vite";
import { tanstackRouter } from "@tanstack/router-plugin/vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

/**
 * The app (React). `bun dev`: Vite on 3300, the Hub server (API) on 3301, /api forwarded to it.
 * `bun run build`: dist/, served by the server itself (`bun run start`).
 */
export default defineConfig({
  server: { port: 3300, proxy: { "/api": { target: "http://127.0.0.1:3301", changeOrigin: false } } },
  resolve: { tsconfigPaths: true, dedupe: ["powerbi-client"] },
  plugins: [tanstackRouter({ target: "react", autoCodeSplitting: true }), react(), tailwindcss()],
});
