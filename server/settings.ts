/**
 * The settings, kept by the server: the .env file (and the environment) gives the defaults, the changes
 * made in Settings are saved in .local/settings.json (readable by the user only) and win over them; a
 * reset forgets the changes of a part, back to the .env values. Secrets (AI keys, Confluence token) stay
 * here and are never sent to the browser, which sees `publicSettings()` only.
 */
import { existsSync, readFileSync } from "node:fs";
import { chmod, mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { SCREENS, type Provider, type PublicSettings, type ScreenId, type SettingsPatch } from "@shared/types";

export const DATA_DIR = process.env.HUB_DATA_DIR ?? join(process.cwd(), ".local");
const FILE = join(DATA_DIR, "settings.json");

/** The .env file of the folder the server runs in (also for the executable): values not already in the environment. */
function loadDotEnv() {
  const file = join(process.cwd(), ".env");
  if (!existsSync(file)) return;
  for (const line of readFileSync(file, "utf8").split(/\r?\n/)) {
    const m = line.match(/^\s*(?:export\s+)?([A-Za-z_][\w]*)\s*=\s*(.*)\s*$/);
    // Bun reads .env too, but takes an unquoted "#…" as a comment (a color comes out empty): read again here.
    if (!m || process.env[m[1]]) continue;
    let v = m[2].trim();
    if (/^(["']).*\1$/.test(v)) v = v.slice(1, -1);
    else v = v.replace(/\s+#.*$/, "");
    if (v.startsWith("#") && !/^#[0-9a-f]{3,8}$/i.test(v)) continue;
    if (v) process.env[m[1]] = v;
  }
}
loadDotEnv();

export type Settings = {
  general: { primaryColor: string; screens: Record<ScreenId, boolean>; mascot: boolean };
  ai: {
    provider: Provider;
    model: string;
    effort: "low" | "medium" | "high";
    maxSteps: number;
    anthropicKey: string;
    maia: { url: string; clientId: string; clientSecret: string; maison: string; useCaseId: string; environment: string };
    claudeCode: { baseUrl: string; headers: string; keyHelper: string; key: string };
  };
  powerbi: { maxRows: number; timeoutMs: number; concurrency: number; cacheMs: number };
  /** Three root folders: projects (a sub-folder each), the reports' catalog and the models' catalog (a page each). */
  confluence: { url: string; email: string; key: string; projectsFolderId: string; reportsFolderId: string; datasetsFolderId: string };
};

/** A folder or page id from what was given: the id itself, or a link (…/folder/123…, …/pages/123…). */
export const folderIdOf = (v: string) => v.match(/\/(?:folder|pages)\/(\d+)/)?.[1] ?? v.trim();
/** A folder id from the environment (an id or a Confluence link), empty when it is not one. */
const envFolder = (name: string) => {
  const id = folderIdOf(env[name] ?? "");
  return /^\d+$/.test(id) ? id : "";
};

const env = process.env;
const num = (v: string | undefined, d: number) => (v && !Number.isNaN(Number(v)) ? Number(v) : d);
/** The defaults: the .env file / environment, else the app's own. */
const defaults = (): Settings => ({
  general: { primaryColor: env.HUB_PRIMARY_COLOR || "#b39251", screens: Object.fromEntries(SCREENS.map((s) => [s, true])) as Record<ScreenId, boolean>, mascot: env.HUB_MASCOT !== "false" },
  ai: {
    provider: (env.HUB_AI_PROVIDER as Provider) || (env.MAIA_CLIENT_ID ? "maia" : "anthropic"),
    model: env.HUB_AI_MODEL || (env.MAIA_CLIENT_ID ? "claude-sonnet-5-5" : "claude-opus-5-5"),
    effort: (env.HUB_AI_EFFORT as Settings["ai"]["effort"]) || "medium",
    maxSteps: num(env.HUB_AI_MAX_STEPS, 40),
    anthropicKey: env.ANTHROPIC_API_KEY ?? env.CLAUDE_API_KEY ?? "",
    maia: { url: env.MAIA_URL || "https://genapi.lvmh.com", clientId: env.MAIA_CLIENT_ID ?? "", clientSecret: env.MAIA_CLIENT_SECRET ?? "", maison: env.MAIA_MAISON || "GRP", useCaseId: env.MAIA_USE_CASE || "hub", environment: env.MAIA_ENVIRONMENT || "PRD" },
    claudeCode: { baseUrl: env.ANTHROPIC_BASE_URL && !/api\.anthropic\.com/.test(env.ANTHROPIC_BASE_URL) ? env.ANTHROPIC_BASE_URL : "", headers: env.ANTHROPIC_CUSTOM_HEADERS ?? "", keyHelper: env.OKTA_SCRIPT ?? "", key: "" },
  },
  powerbi: { maxRows: num(env.HUB_DAX_MAX_ROWS, 2000), timeoutMs: num(env.HUB_DAX_TIMEOUT_MS, 20_000), concurrency: num(env.HUB_DAX_CONCURRENCY, 2), cacheMs: num(env.HUB_DAX_CACHE_MS, 60_000) },
  confluence: { url: env.CONFLUENCE_API_URL ?? "", email: env.CONFLUENCE_EMAIL ?? "", key: env.CONFLUENCE_API_KEY ?? "", projectsFolderId: envFolder("CONFLUENCE_PROJECTS_FOLDER_ID"), reportsFolderId: envFolder("CONFLUENCE_REPORT_FOLDER_ID"), datasetsFolderId: envFolder("CONFLUENCE_DATASET_FOLDER_ID") },
});

/** Settings a .env value gives (shown in Settings: "from .env"). */
const ENV_KEYS: Record<string, string[]> = {
  "general.primaryColor": ["HUB_PRIMARY_COLOR"],
  "ai.provider": ["HUB_AI_PROVIDER"],
  "ai.model": ["HUB_AI_MODEL"],
  "ai.anthropicKey": ["ANTHROPIC_API_KEY", "CLAUDE_API_KEY"],
  "ai.maia.clientId": ["MAIA_CLIENT_ID"],
  "ai.maia.clientSecret": ["MAIA_CLIENT_SECRET"],
  "ai.maia.url": ["MAIA_URL"],
  "ai.maia.maison": ["MAIA_MAISON"],
  "ai.maia.useCaseId": ["MAIA_USE_CASE"],
  "ai.maia.environment": ["MAIA_ENVIRONMENT"],
  "ai.claudeCode.baseUrl": ["ANTHROPIC_BASE_URL"],
  "ai.claudeCode.keyHelper": ["OKTA_SCRIPT"],
  "powerbi.maxRows": ["HUB_DAX_MAX_ROWS"],
  "confluence.url": ["CONFLUENCE_API_URL"],
  "confluence.email": ["CONFLUENCE_EMAIL"],
  "confluence.key": ["CONFLUENCE_API_KEY"],
  "confluence.projectsFolderId": ["CONFLUENCE_PROJECTS_FOLDER_ID"],
  "confluence.reportsFolderId": ["CONFLUENCE_REPORT_FOLDER_ID"],
  "confluence.datasetsFolderId": ["CONFLUENCE_DATASET_FOLDER_ID"],
};

/** Only the changes made in Settings (merged over the defaults at each start). */
let saved: Record<string, any> | undefined;
let current: Settings | undefined;

/** Deep merge of plain objects (arrays and values replaced). */
function merge<T>(base: T, patch: unknown): T {
  if (!patch || typeof patch !== "object" || Array.isArray(patch)) return (patch === undefined ? base : patch) as T;
  const out: any = { ...(base as any) };
  for (const [k, v] of Object.entries(patch)) if (v !== undefined) out[k] = merge((base as any)?.[k], v);
  return out;
}

export async function settings(): Promise<Settings> {
  if (!current) {
    saved = await readFile(FILE, "utf8").then(JSON.parse).catch(() => ({}));
    current = merge(defaults(), saved);
  }
  return current!;
}

async function save() {
  await mkdir(DATA_DIR, { recursive: true });
  await writeFile(FILE, JSON.stringify(saved, null, 2));
  await chmod(FILE, 0o600).catch(() => undefined);
  current = merge(defaults(), saved);
}

/** Whether the reports' and models' catalog can be changed from the app (CONFLUENCE_CATALOG_EDIT). */
export const catalogEditable = () => /^(1|true|yes)$/i.test(env.CONFLUENCE_CATALOG_EDIT ?? "");

export async function publicSettings(): Promise<PublicSettings> {
  const s = await settings();
  return {
    general: s.general,
    ai: {
      provider: s.ai.provider,
      model: s.ai.model,
      effort: s.ai.effort,
      maxSteps: s.ai.maxSteps,
      anthropic: { keySet: !!s.ai.anthropicKey },
      maia: { url: s.ai.maia.url, clientIdSet: !!s.ai.maia.clientId, clientSecretSet: !!s.ai.maia.clientSecret, maison: s.ai.maia.maison, useCaseId: s.ai.maia.useCaseId, environment: s.ai.maia.environment },
      claudeCode: { baseUrl: s.ai.claudeCode.baseUrl, headers: s.ai.claudeCode.headers, keyHelper: s.ai.claudeCode.keyHelper, keySet: !!s.ai.claudeCode.key },
    },
    powerbi: s.powerbi,
    confluence: { url: s.confluence.url, email: s.confluence.email, keySet: !!s.confluence.key, projectsFolderId: s.confluence.projectsFolderId, reportsFolderId: s.confluence.reportsFolderId, datasetsFolderId: s.confluence.datasetsFolderId, catalogEdit: catalogEditable() },
    env: Object.entries(ENV_KEYS).filter(([, keys]) => keys.some((k) => env[k])).map(([path]) => path),
  };
}

/** Applies a change from Settings; an empty secret leaves the current one as it is. */
export async function updateSettings(patch: SettingsPatch) {
  await settings();
  const { anthropicKey, maia, claudeCode, ...ai } = patch.ai ?? {};
  const filled = <T extends Record<string, unknown>>(o: T | undefined) => (o ? Object.fromEntries(Object.entries(o).filter(([k, v]) => v !== undefined && !(v === "" && /key|secret|clientId/i.test(k)))) : undefined);
  const c = patch.confluence;
  saved = merge(saved, {
    general: patch.general,
    ai: { ...ai, ...(anthropicKey ? { anthropicKey } : {}), maia: filled(maia), claudeCode: filled(claudeCode) },
    powerbi: patch.powerbi,
    confluence: c && filled(Object.fromEntries(Object.entries(c).map(([k, v]) => [k, /FolderId$/.test(k) && typeof v === "string" ? folderIdOf(v) : v]))),
  });
  await save();
  return publicSettings();
}

/** Forgets the changes of a part of the settings (general, ai, powerbi, confluence): back to the .env values. */
export async function resetSettings(part: "general" | "ai" | "powerbi" | "confluence") {
  await settings();
  delete saved![part];
  await save();
  return publicSettings();
}
