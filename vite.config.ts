import { readFileSync } from "node:fs";
import { defineConfig, type Plugin } from "vite";
import { tanstackRouter } from "@tanstack/router-plugin/vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { viteSingleFile } from "vite-plugin-singlefile";

/**
 * The page's content security policy (docs/security.md), written into index.html: scripts only from
 * the app itself, network calls only to the local server and the AI API (its origin read from
 * src/agent/llm.ts) — the Microsoft APIs are called by the local server —, frames only from Power BI
 * and Office. A script slipped into the page could not load code from elsewhere nor send data to
 * another site. HUB_CSP_CONNECT adds origins (a local server on another machine); HUB_NO_CSP=1 leaves it out.
 */
function contentSecurityPolicy(): Plugin {
  const ai = readFileSync("src/agent/llm.ts", "utf8").match(/AI_BASE_URL\s*=\s*"(https?:\/\/[^/"]+)/)?.[1] ?? "";
  const policy = [
    "default-src 'self'",
    // Inline: Vite's dev preamble, and the whole app in the single-file export. blob: the AI's sandboxed worker.
    "script-src 'self' 'unsafe-inline' blob:",
    "worker-src 'self' blob:",
    "style-src 'self' 'unsafe-inline'",
    "font-src 'self' data:",
    "img-src 'self' data: blob: https:",
    `connect-src 'self' http://localhost:* http://127.0.0.1:* ws://localhost:* ws://127.0.0.1:* ${ai} ${process.env.HUB_CSP_CONNECT ?? ""}`.trim(),
    "frame-src https://*.powerbi.com https://*.analysis.windows.net https://*.pbidedicated.windows.net https://*.sharepoint.com https://*.officeapps.live.com https://*.office.com https://*.microsoftonline.com https://*.microsoft.com",
    "object-src 'none'",
    "base-uri 'self'",
    "form-action 'self'",
  ].join("; ");
  return {
    name: "hub-csp",
    transformIndexHtml: () => (process.env.HUB_NO_CSP === "1" ? [] : [{ tag: "meta", attrs: { "http-equiv": "Content-Security-Policy", content: policy }, injectTo: "head-prepend" }, { tag: "meta", attrs: { name: "referrer", content: "no-referrer" }, injectTo: "head-prepend" }]),
  };
}

/**
 * The app is a single page (no server of its own): `bun dev` serves it with the local token server
 * beside it. `--mode html` builds it into one standalone HTML file (dist/hub.html, routes in the URL hash).
 */
export default defineConfig(({ mode }) => ({
  // Only this computer reaches the dev server (never the local network).
  server: { port: 3300, host: "localhost", strictPort: true },
  // One powerbi-client: powerbi-report-authoring adds its methods (getCapabilities…) to the copy it imports.
  resolve: { tsconfigPaths: true, dedupe: ["powerbi-client"] },
  base: mode === "html" ? "./" : "/",
  build: mode === "html" ? { outDir: "dist/html", emptyOutDir: true, assetsInlineLimit: 100_000_000, cssCodeSplit: false } : undefined,
  plugins: [contentSecurityPolicy(), tanstackRouter({ target: "react", autoCodeSplitting: false }), react(), tailwindcss(), ...(mode === "html" ? [viteSingleFile()] : [])],
}));
