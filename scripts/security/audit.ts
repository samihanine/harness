/**
 * bun run audit: security check of the project.
 * 1. Dependencies: known vulnerabilities of the installed packages (bun audit).
 * 2. The repository: secrets or local files committed by mistake (keys, tokens, .env, .local).
 * Exits with an error when something is found (usable in CI).
 */
import { ROOT, SEVERITIES, audit, count } from "./lib";

let problems = 0;

console.log("\n1. Dependencies (bun audit)");
const found = audit();
for (const [name, list] of Object.entries(found)) {
  console.log(`\n  ${name}`);
  for (const a of list.sort((x, y) => SEVERITIES.indexOf(y.severity) - SEVERITIES.indexOf(x.severity))) console.log(`    ${a.severity.padEnd(8)} ${a.title}\n             ${a.vulnerable_versions} — ${a.url}`);
}
const total = Object.values(found).flat().length;
problems += total;
console.log(total ? `\n  ${total} vulnerabilities (${count(found)}). Fix: bun run audit:fix` : "  No known vulnerability.");

console.log("\n2. Repository (committed secrets and local files)");
/** Files that must never be committed. */
const FORBIDDEN = [/(^|\/)\.env(\.|$)(?!example)/, /(^|\/)\.local\//, /tokens\.json$/, /\.(pem|key|p12|pfx)$/, /(^|\/)id_(rsa|ed25519)$/];
/** Secrets in a file's text. */
const SECRETS: [string, RegExp][] = [
  ["private key", /-----BEGIN [A-Z ]*PRIVATE KEY-----/],
  ["AWS key", /\bAKIA[0-9A-Z]{16}\b/],
  ["GitHub token", /\bgh[pousr]_[A-Za-z0-9]{36,}\b/],
  ["OpenAI-style key", /\bsk-[A-Za-z0-9_-]{32,}\b/],
  ["Slack token", /\bxox[abprs]-[A-Za-z0-9-]{10,}\b/],
  ["JWT / access token", /\beyJ[A-Za-z0-9_-]{20,}\.eyJ[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{20,}/],
  ["client secret", /client_secret\s*[:=]\s*["'][^"'\s]{16,}["']/i],
  ["password", /\bpassword\s*[:=]\s*["'][^"'\s]{8,}["']/i],
];
const files = Bun.spawnSync(["git", "ls-files"], { cwd: ROOT, stdout: "pipe" }).stdout.toString().split("\n").filter(Boolean);
const issues: string[] = [];
for (const f of files) {
  if (FORBIDDEN.some((r) => r.test(f))) issues.push(`${f}: must not be committed (local file or secret)`);
  if (f === "bun.lock" || /\.(png|jpe?g|gif|webp|ico|woff2?|ttf|zip|pdf|pptx|xlsx)$/i.test(f)) continue;
  const text = await Bun.file(ROOT + f).text().catch(() => "");
  for (const [kind, r] of SECRETS) {
    const at = text.search(r);
    if (at >= 0) issues.push(`${f}:${text.slice(0, at).split("\n").length}: looks like a ${kind}`);
  }
}
problems += issues.length;
console.log(issues.length ? issues.map((i) => `  ${i}`).join("\n") : `  ${files.length} files checked: nothing found.`);

console.log(problems ? `\n✗ ${problems} problem(s).` : "\n✓ No security problem found.");
process.exit(problems ? 1 : 0);
