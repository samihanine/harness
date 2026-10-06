import { getRequestListener } from "@hono/node-server";
import tailwindcss from "@tailwindcss/vite";
import { tanstackRouter } from "@tanstack/router-plugin/vite";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import type { Plugin } from "vite";

/** Dev: the API (server/app.ts) runs inside the Vite server, so `bun dev` is the only command. */
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
  base: "/",
  plugins: [tanstackRouter({ target: "react", autoCodeSplitting: true }), react(), tailwindcss(), api()],
  resolve: { tsconfigPaths: true },
  server: { port: 3201, strictPort: true, host: "127.0.0.1" },
});
