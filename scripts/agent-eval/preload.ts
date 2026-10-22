/**
 * Runs the app's agents outside the browser: the IndexedDB store becomes an in-memory one (the AI key
 * comes from .env / .env.local). Microsoft calls fail clearly, unless EVAL_LIVE=1: then the real
 * src/server/ms.ts runs here (same tokens as the dev server, .local/tokens.json), its server
 * functions called directly. Modules are swapped when loaded (Bun's runtime plugins do not see
 * tsconfig aliases).
 */
import { plugin } from "bun";
import { readFileSync } from "node:fs";

const fake = (name: string) => readFileSync(new URL(`./${name}`, import.meta.url), "utf8");
/** createServerFn(...).inputValidator(...).handler(fn) → fn itself (called as fn({ data })). */
const SHIM = `const createServerFn = (_?: unknown) => { const b: any = { inputValidator: () => b, middleware: () => b, handler: (fn: any) => (arg?: any) => fn(arg ?? {}) }; return b; };`;

plugin({
  name: "agent-eval",
  setup(build) {
    build.onLoad({ filter: /\/src\/lib\/store\.ts$/ }, () => ({ contents: fake("fake-store.ts"), loader: "ts" }));
    build.onLoad({ filter: /\/src\/server\/ms\.ts$/ }, (args) => ({
      contents: process.env.EVAL_LIVE === "1" ? readFileSync(args.path, "utf8").replace(/import \{ createServerFn \} from "@tanstack\/react-start";/, SHIM) : fake("fake-server.ts"),
      loader: "ts",
    }));
  },
});
