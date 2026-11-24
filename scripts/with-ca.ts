/**
 * Runs a command with NODE_EXTRA_CA_CERTS=./ca.crt when that file is at the root of the repo (a company
 * certificate, e.g. a TLS-inspecting proxy), and runs it unchanged when it is not there: no error either way.
 *   bun scripts/with-ca.ts <command> [args…]
 */
import { existsSync } from "node:fs";
import { join } from "node:path";

const ca = join(import.meta.dir, "..", "ca.crt");
const env = { ...process.env, ...(existsSync(ca) && !process.env.NODE_EXTRA_CA_CERTS ? { NODE_EXTRA_CA_CERTS: ca } : {}) };
const [cmd, ...args] = process.argv.slice(2);
if (!cmd) {
  console.error("Usage: bun scripts/with-ca.ts <command> [args…]");
  process.exit(1);
}
const child = Bun.spawn([cmd, ...args], { env, stdio: ["inherit", "inherit", "inherit"] });
for (const sig of ["SIGINT", "SIGTERM"] as const) process.on(sig, () => child.kill(sig));
process.exit(await child.exited);
