/**
 * bun run audit:fix: fixes the dependencies' known vulnerabilities without breaking the app.
 * 1. Compatible updates (bun update, within the ranges of package.json).
 * 2. Each package still vulnerable goes to the lowest version above the installed one without the
 *    vulnerability (the same major when there is one): a direct dependency in package.json, a
 *    dependency of a dependency through "overrides".
 * 3. Typecheck + build: if they fail, everything is put back as it was.
 */
import { readFileSync } from "node:fs";
import { ROOT, audit, backup, check, count, pkg, run, savePkg, type Advisory } from "./lib";

const before = audit();
if (!Object.keys(before).length) {
  console.log("✓ No known vulnerability: nothing to fix.");
  process.exit(0);
}
console.log(`Vulnerabilities: ${count(before)}.`);
const saved = backup();

console.log("\n1. Compatible updates (bun update)…");
run(["bun", "update"]);

/** The versions of a package in bun.lock (several when dependencies need different ones; aliases too). */
function locked(name: string) {
  const text = readFileSync(`${ROOT}bun.lock`, "utf8");
  const escaped = name.replace(/[.*+?^${}()|[\]\\/]/g, "\\$&");
  return [...text.matchAll(new RegExp(`\\["${escaped}@([^"]+)"`, "g"))].map((m) => m[1]);
}
const major = (v: string) => v.split(".")[0];

type Meta = { versions: Record<string, { dependencies?: Record<string, string> }>; "dist-tags": { latest: string } };
const metas = new Map<string, Meta>();
async function meta(name: string) {
  if (!metas.has(name)) {
    const res = await fetch(`https://registry.npmjs.org/${name.replace("/", "%2F")}`);
    if (!res.ok) throw new Error(`${name}: registry ${res.status}`);
    metas.set(name, (await res.json()) as Meta);
  }
  return metas.get(name)!;
}

/**
 * The smallest move out of every vulnerable range: the lowest published version above the installed
 * (vulnerable) one that is in none of them, in the same major when there is one (fewer breaking changes).
 */
async function safeVersion(name: string, advisories: Advisory[]) {
  const m = await meta(name);
  const vulnerable = locked(name).filter((v) => advisories.some((a) => Bun.semver.satisfies(v, a.vulnerable_versions)));
  const from = vulnerable.sort((a, b) => Bun.semver.order(a, b))[0] ?? locked(name)[0] ?? "0.0.0";
  const safe = Object.keys(m.versions)
    .filter((v) => !v.includes("-") && Bun.semver.order(v, from) > 0 && Bun.semver.order(v, m["dist-tags"].latest) <= 0)
    .filter((v) => advisories.every((a) => !Bun.semver.satisfies(v, a.vulnerable_versions)))
    .sort((a, b) => Bun.semver.order(a, b));
  return { from, to: safe.find((v) => major(v) === major(from)) ?? safe[0] };
}

const DIRECT = ["dependencies", "devDependencies"] as const;
const fieldOf = (p: ReturnType<typeof pkg>, name: string) => DIRECT.find((f) => p[f]?.[name]);
/** Sets a direct dependency's version in the same style as before (an exact version stays exact, a range keeps its "^"). */
function setVersion(p: ReturnType<typeof pkg>, name: string, version: string) {
  const f = fieldOf(p, name)!;
  p[f]![name] = /^\d/.test(p[f]![name]) ? version : `^${version}`;
}

let left = audit();
const p0 = pkg();
const direct = Object.entries(left).filter(([name]) => fieldOf(p0, name));
if (direct.length) {
  console.log("\n2. Vulnerable direct dependencies → the smallest safe version");
  const next = pkg();
  for (const [name, advisories] of direct) {
    const { from, to } = await safeVersion(name, advisories);
    if (!to) {
      console.log(`  ${name}: no published version without the vulnerability yet (left as is)`);
      continue;
    }
    console.log(`  ${name}: ${from} → ${to}`);
    setVersion(next, name, to);
    // Packages released together (e.g. @tanstack/*): our other direct dependencies it needs at an exact version move with it.
    for (const [dep, range] of Object.entries((await meta(name)).versions[to]?.dependencies ?? {})) {
      if (fieldOf(next, dep) && /^\d/.test(range) && next[fieldOf(next, dep)!]![dep].replace(/^\^/, "") !== range) {
        console.log(`    with ${dep} → ${range} (needed by ${name} ${to})`);
        setVersion(next, dep, range);
      }
    }
  }
  savePkg(next);
  run(["bun", "install"]);
  left = audit();
}

// What is still vulnerable is a copy installed by another package (also a direct one's own copy, e.g. an older zod under a library).
const indirect = Object.entries(left);
if (indirect.length) {
  console.log("\n2b. Vulnerable copies installed by other packages → overrides");
  const next = pkg();
  for (const [name, advisories] of indirect) {
    const { from, to } = await safeVersion(name, advisories);
    if (!to) {
      console.log(`  ${name}: no published version without the vulnerability yet (left as is)`);
      continue;
    }
    console.log(`  ${name}: ${from} → ${to}`);
    next.overrides = { ...next.overrides, [name]: `^${to}` };
  }
  savePkg(next);
  run(["bun", "install"]);
}

console.log("\n3. Checking that the app still works…");
if (!check()) {
  console.log("\n✗ The fixes break the app: everything is put back as it was. Fix by hand (see the errors above), or run bun run update:libs.");
  saved.restore();
  process.exit(1);
}
const after = audit();
console.log(Object.keys(after).length ? `\nLeft: ${count(after)} (${Object.keys(after).join(", ")}): no fix published yet, or a fix that needs code changes.` : "\n✓ No known vulnerability left.");
process.exit(Object.keys(after).length ? 1 : 0);
