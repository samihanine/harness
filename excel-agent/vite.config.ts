import { getRequestListener } from "@hono/node-server";
import tailwindcss from "@tailwindcss/vite";
import { tanstackRouter } from "@tanstack/router-plugin/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import type { Plugin } from "vite";
import { viteSingleFile } from "vite-plugin-singlefile";

/** SINGLE_FILE=1 builds one self-contained HTML (bun run export-html; local files only, no sign-in API). */
const singleFile = process.env.SINGLE_FILE === "1";

/** Dev: the API (server/app.ts, Microsoft sign-in) runs inside the Vite server: `bun dev` is the only command. */
const api = (): Plugin => ({
  name: "api",
  configureServer(server) {
    server.middlewares.use(async (req, res, next) => {
      if (!req.url?.startsWith("/api/")) return next();
      const { app } = (await server.ssrLoadModule("/server/app.ts")) as typeof import("./server/app");
      return getRequestListener(app.fetch)(req, res);
    });
  },
});

export default defineConfig({
  base: singleFile ? "./" : "/",
  plugins: [
    tanstackRouter({ target: "react", autoCodeSplitting: !singleFile }),
    react(),
    tailwindcss(),
    api(),
    ...(singleFile ? [viteSingleFile({ removeViteModuleLoader: true })] : []),
  ],
  resolve: { tsconfigPaths: true },
  server: { port: 3202, strictPort: true },
});
