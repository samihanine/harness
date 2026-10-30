/**
 * bun run update:libs: every dependency to its latest version, without breaking the app.
 * All at once first; if the typecheck or the build fails, back to how it was, then one
 * family at a time (packages released together move together): each update is kept only when the
 * app still builds (the others are listed with why).
 * bun run update:libs --dry: only lists the outdated libraries.
 */
import { backup, check, pkg, run } from "./lib";

// --dry: only list what would be updated.
if (process.argv.includes("--dry")) {
  run(["bun", "outdated"], false);
  process.exit(0);
}

const all = backup();
console.log("All dependencies to their latest version (bun update --latest)…");
run(["bun", "update", "--latest"]);
if (check()) {
  console.log("\n✓ Everything updated and the app still builds.");
  process.exit(0);
}
console.log("\nSomething breaks with everything updated: back to how it was, then one package at a time.\n");
all.restore();

const p = pkg();
const dev = new Set(Object.keys(p.devDependencies ?? {}));
const version = (q: ReturnType<typeof pkg>, name: string) => (dev.has(name) ? q.devDependencies : q.dependencies)?.[name];
/**
 * Packages released together move together (all or none): a family (@tiptap/*, @tanstack/*…), and
 * react with react-dom. Two versions of one family side by side break at run time, not at build.
 */
const family = (name: string) => (/^(react|react-dom)$/.test(name) ? "react" : name.startsWith("@") && !name.startsWith("@types/") ? name.split("/")[0] : name);
const groups = new Map<string, string[]>();
for (const name of [...Object.keys(p.dependencies ?? {}), ...dev]) groups.set(family(name), [...(groups.get(family(name)) ?? []), name]);
const kept: string[] = [];
const skipped: string[] = [];
for (const names of groups.values()) {
  const one = backup();
  for (const d of [false, true]) {
    const list = names.filter((n) => dev.has(n) === d);
    if (list.length) run(["bun", "add", ...(d ? ["-d"] : []), ...list.map((n) => `${n}@latest`)]);
  }
  const moved = names.filter((n) => version(pkg(), n) !== version(p, n));
  if (!moved.length) continue; // already the latest
  const label = moved.map((n) => `${n} ${version(p, n)} → ${version(pkg(), n)}`).join(", ");
  console.log(label);
  if (check("  ")) kept.push(...moved.map((n) => `${n} ${version(pkg(), n)}`));
  else {
    skipped.push(...moved.map((n) => `${n} (stays ${version(p, n)})`));
    one.restore();
  }
}
console.log(`\nUpdated: ${kept.join(", ") || "none"}`);
console.log(`Not updated (they break the typecheck or the build: needs code changes): ${skipped.join(", ") || "none"}`);
