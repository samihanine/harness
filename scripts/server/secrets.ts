/**
 * The secrets the local server keeps (the Microsoft refresh token, the Confluence key): in the
 * operating system's credential store (Windows Credential Manager, macOS Keychain, libsecret) through
 * Bun.secrets; where there is none, in .local/secrets.json readable by the user only. Never sent to
 * the browser, never logged.
 */
import { chmod, mkdir, readFile, rm, writeFile } from "node:fs/promises";

const SERVICE = "hub-local-server";
const FILE = ".local/secrets.json";
type Store = { get(o: { service: string; name: string }): Promise<string | null>; set(o: { service: string; name: string; value: string }): Promise<void>; delete(o: { service: string; name: string }): Promise<boolean> };
const os = (globalThis as { Bun?: { secrets?: Store } }).Bun?.secrets;
/** Decided at the first use: the OS store when it works on this machine. */
let useOs: Promise<boolean> | undefined;
const osWorks = () =>
  (useOs ??= (async () => {
    if (!os || process.env.HUB_SECRETS === "file") return false;
    try {
      await os.get({ service: SERVICE, name: "probe" });
      return true;
    } catch {
      return false;
    }
  })());

const readFileStore = async (): Promise<Record<string, string>> => readFile(FILE, "utf8").then(JSON.parse, () => ({}));
async function writeFileStore(values: Record<string, string>) {
  await mkdir(".local", { recursive: true });
  await writeFile(FILE, JSON.stringify(values), { mode: 0o600 });
  await chmod(FILE, 0o600).catch(() => undefined);
}

export async function getSecret(name: string): Promise<string | undefined> {
  if (await osWorks()) return (await os!.get({ service: SERVICE, name })) ?? undefined;
  return (await readFileStore())[name];
}

export async function setSecret(name: string, value: string | undefined) {
  if (await osWorks()) return void (value === undefined ? await os!.delete({ service: SERVICE, name }) : await os!.set({ service: SERVICE, name, value }));
  const all = await readFileStore();
  if (value === undefined) delete all[name];
  else all[name] = value;
  await writeFileStore(all);
}

/** The refresh token of earlier versions (.local/tokens.json, in clear): moved to the store, the file removed. */
export async function migrateOldTokens() {
  const old = await readFile(".local/tokens.json", "utf8").then(JSON.parse, () => undefined);
  if (old?.refreshToken && !(await getSecret("session"))) await setSecret("session", JSON.stringify({ refreshToken: old.refreshToken, account: old.account }));
  await rm(".local/tokens.json", { force: true });
  // Access tokens written by an earlier "login" command.
  await rm(".local/access-tokens.json", { force: true });
}

export const secretStore = async () => ((await osWorks()) ? "the system's credential store" : `${FILE} (user-only file)`);
