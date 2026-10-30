/**
 * Runs the app's agents outside the browser: the IndexedDB store becomes an in-memory one (the AI key
 * comes from .env / .env.local). Microsoft calls fail clearly, unless EVAL_LIVE=1: then the real
 * src/services/ms.ts runs here, its tokens read from .local/access-tokens.json (`bun run token:login`) or
 * asked to the token server. Modules are swapped when loaded (Bun's runtime plugins do not see
 * tsconfig aliases).
 */
import { plugin } from "bun";
import { existsSync, readFileSync } from "node:fs";

const fake = (name: string) => readFileSync(new URL(`./${name}`, import.meta.url), "utf8");

// The browser's localStorage (tokens, options), in memory; seeded with the tokens of the last sign-in.
const memory = new Map<string, string>();
Object.assign(globalThis, {
  localStorage: { getItem: (k: string) => memory.get(k) ?? null, setItem: (k: string, v: string) => void memory.set(k, String(v)), removeItem: (k: string) => void memory.delete(k), clear: () => memory.clear(), key: (i: number) => [...memory.keys()][i] ?? null, get length() { return memory.size; } },
});
if (process.env.EVAL_LIVE === "1" && existsSync(".local/access-tokens.json")) {
  const saved = JSON.parse(readFileSync(".local/access-tokens.json", "utf8")) as Record<string, { token: string; expiresAt: number; account?: string } | string>;
  for (const [service, t] of Object.entries(saved)) if (t && typeof t === "object" && t.expiresAt > Date.now() + 60_000) memory.set(`ms-token:${service}`, JSON.stringify(t));
}

plugin({
  name: "agent-eval",
  setup(build) {
    build.onLoad({ filter: /\/src\/lib\/store\.ts$/ }, () => ({ contents: fake("fake-store.ts"), loader: "ts" }));
    if (process.env.EVAL_LIVE !== "1") build.onLoad({ filter: /\/src\/services\/ms\.ts$/ }, () => ({ contents: fake("fake-server.ts"), loader: "ts" }));
  },
});
