/**
 * Runs the app's agents outside the browser: the IndexedDB store becomes an in-memory one (the AI key
 * comes from .env / .env.local). Microsoft calls fail clearly, unless EVAL_LIVE=1: then the real
 * src/services/ms.ts runs here and calls Microsoft through the local server (start it: `bun run token`,
 * signed in: `bun run token:login`). Modules are swapped when loaded (Bun's runtime plugins do not see
 * tsconfig aliases).
 */
import { plugin } from "bun";
import { readFileSync } from "node:fs";

const fake = (name: string) => readFileSync(new URL(`./${name}`, import.meta.url), "utf8");

// The browser's localStorage (options), in memory.
const memory = new Map<string, string>();
Object.assign(globalThis, {
  localStorage: { getItem: (k: string) => memory.get(k) ?? null, setItem: (k: string, v: string) => void memory.set(k, String(v)), removeItem: (k: string) => void memory.delete(k), clear: () => memory.clear(), key: (i: number) => [...memory.keys()][i] ?? null, get length() { return memory.size; } },
});
plugin({
  name: "agent-eval",
  setup(build) {
    build.onLoad({ filter: /\/src\/lib\/store\.ts$/ }, () => ({ contents: fake("fake-store.ts"), loader: "ts" }));
    if (process.env.EVAL_LIVE !== "1") build.onLoad({ filter: /\/src\/services\/ms\.ts$/ }, () => ({ contents: fake("fake-server.ts"), loader: "ts" }));
  },
});
