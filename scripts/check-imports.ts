/**
 * The boundaries of the code (docs/structure.md), checked with the types (bun run typecheck):
 * - the shared code (lib, services, agent, components, prompts, config) never imports a feature, the app layer or a route;
 * - a feature never imports another feature (what they share lives in the shared code);
 * - only the app layer (src/app, src/routes) knows every feature.
 */
import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";

const SRC = resolve(import.meta.dir, "../src");
const files = (dir: string): string[] => readdirSync(dir).flatMap((f) => (statSync(join(dir, f)).isDirectory() ? files(join(dir, f)) : /\.tsx?$/.test(f) ? [join(dir, f)] : []));
/** Where an import leads, as a path under src ("features/excel/table"), or undefined for a package. */
const target = (file: string, spec: string) => (spec.startsWith("@/") ? spec.slice(2) : spec.startsWith(".") ? relative(SRC, resolve(dirname(file), spec)) : undefined);
const area = (path: string) => {
  const [top, name] = path.split("/");
  return top === "features" ? `features/${name}` : top.replace(/\.tsx?$/, "");
};
const SHARED = ["lib", "services", "agent", "components", "prompts", "config"];

const problems: string[] = [];
for (const file of files(SRC)) {
  const from = area(relative(SRC, file));
  if (from === "routeTree.gen") continue;
  for (const m of readFileSync(file, "utf8").matchAll(/(?:import|export)[^"']*?from\s+["']([^"']+)["']|import\(\s*["']([^"']+)["']\s*\)/g)) {
    const to = target(file, m[1] ?? m[2]);
    if (!to) continue;
    const dest = area(to);
    const bad = (SHARED.includes(from) && (dest.startsWith("features/") || dest === "app" || dest === "routes")) || (from.startsWith("features/") && ((dest.startsWith("features/") && dest !== from) || dest === "app" || dest === "routes"));
    if (bad) problems.push(`${relative(SRC, file)} → ${m[1] ?? m[2]}`);
  }
}
if (problems.length) {
  console.error(`Imports across boundaries (docs/structure.md):\n${problems.map((p) => `  ${p}`).join("\n")}`);
  process.exit(1);
}
console.log("Boundaries: ok");
