/**
 * Shared by the security scripts: running commands, the project's checks (typecheck + build: what a
 * dependency change must not break), a backup of package.json / bun.lock to roll back, and the audit.
 */
import { copyFileSync, readFileSync, writeFileSync } from "node:fs";

export const ROOT = new URL("../../", import.meta.url).pathname;
const FILES = ["package.json", "bun.lock"];

/** Runs a command in the project; `quiet`: its output is kept and shown only when it fails. */
export function run(cmd: string[], quiet = true) {
  const p = Bun.spawnSync(cmd, { cwd: ROOT, stdout: quiet ? "pipe" : "inherit", stderr: quiet ? "pipe" : "inherit", env: { ...process.env, FORCE_COLOR: "0" } });
  const output = quiet ? `${p.stdout.toString()}${p.stderr.toString()}` : "";
  return { ok: p.exitCode === 0, output };
}

/** What a change of the dependencies must keep working: the types and the production build. */
export function check(label = "Checking") {
  for (const [name, cmd] of [["typecheck", ["bun", "run", "typecheck"]], ["build", ["bun", "run", "build"]]] as const) {
    process.stdout.write(`  ${label}: ${name}… `);
    const r = run([...cmd]);
    if (!r.ok) {
      console.log("FAILED");
      console.log(r.output.split("\n").filter((l) => /error|Error/.test(l)).slice(0, 12).join("\n"));
      return false;
    }
    console.log("ok");
  }
  return true;
}

/** package.json and bun.lock as they are now; `restore` puts them back and reinstalls. */
export function backup() {
  const saved = FILES.map((f) => [f, readFileSync(ROOT + f)] as const);
  return {
    restore() {
      for (const [f, data] of saved) writeFileSync(ROOT + f, data);
      run(["bun", "install"]);
    },
  };
}
export const copy = (from: string, to: string) => copyFileSync(ROOT + from, ROOT + to);

export type Advisory = { id: number; url: string; title: string; severity: "low" | "moderate" | "high" | "critical"; vulnerable_versions: string };
/** The known vulnerabilities of the installed packages, by package. */
export function audit(): Record<string, Advisory[]> {
  const r = Bun.spawnSync(["bun", "audit", "--json"], { cwd: ROOT, stdout: "pipe", stderr: "pipe" });
  const text = r.stdout.toString().trim();
  try {
    return text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`bun audit gave no JSON:\n${text}\n${r.stderr.toString()}`);
  }
}

export const SEVERITIES = ["low", "moderate", "high", "critical"] as const;
export const count = (found: Record<string, Advisory[]>) => {
  const list = Object.values(found).flat();
  return SEVERITIES.map((s) => `${list.filter((a) => a.severity === s).length} ${s}`).join(", ");
};

export const pkg = () => JSON.parse(readFileSync(ROOT + "package.json", "utf8")) as { dependencies?: Record<string, string>; devDependencies?: Record<string, string>; overrides?: Record<string, string> };
export const savePkg = (p: ReturnType<typeof pkg>) => writeFileSync(ROOT + "package.json", `${JSON.stringify(p, null, 2)}\n`);
