/**
 * Runs the app's agents outside the browser: the IndexedDB store becomes an in-memory one (the AI key
 * comes from .env.local) and Microsoft calls fail clearly (the eval never touches the tenant).
 * The modules' contents are swapped when loaded (Bun's runtime plugins do not see tsconfig aliases).
 */
import { plugin } from "bun";
import { readFileSync } from "node:fs";

const fake = (name: string) => readFileSync(new URL(`./${name}`, import.meta.url), "utf8");

plugin({
  name: "agent-eval",
  setup(build) {
    build.onLoad({ filter: /\/src\/lib\/store\.ts$/ }, () => ({ contents: fake("fake-store.ts"), loader: "ts" }));
    build.onLoad({ filter: /\/src\/server\/ms\.ts$/ }, () => ({ contents: fake("fake-server.ts"), loader: "ts" }));
  },
});
