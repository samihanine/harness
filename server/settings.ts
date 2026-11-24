/**
 * The settings, kept by the server in .local/settings.json (readable by the user only): secrets (AI keys,
 * Confluence token) stay here and are never sent to the browser, which sees `publicSettings()` only.
 * Environment variables give the first values (CLAUDE_API_KEY / ANTHROPIC_API_KEY, MAIA_*, CONFLUENCE_*).
 */
import { chmod, mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import type { Provider, PublicSettings, ScreenId, SettingsPatch } from "@shared/types";
import { SCREENS } from "@shared/types";

export const DATA_DIR = process.env.HUB_DATA_DIR ?? join(process.cwd(), ".local");
const FILE = join(DATA_DIR, "settings.json");

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
  confluence: { url: string; email: string; folderId: string; key: string };
};

const env = process.env;
const defaults = (): Settings => ({
  general: { primaryColor: "#d97757", screens: Object.fromEntries(SCREENS.map((s) => [s, true])) as Record<ScreenId, boolean>, mascot: true },
  ai: {
    provider: env.MAIA_CLIENT_ID ? "maia" : "anthropic",
    model: env.MAIA_CLIENT_ID ? "claude-sonnet-5-5" : "claude-opus-5-5",
    effort: "medium",
    maxSteps: 40,
    anthropicKey: env.ANTHROPIC_API_KEY ?? env.CLAUDE_API_KEY ?? "",
    maia: { url: env.MAIA_URL ?? "https://genapi.lvmh.com", clientId: env.MAIA_CLIENT_ID ?? "", clientSecret: env.MAIA_CLIENT_SECRET ?? "", maison: env.MAIA_MAISON ?? "GRP", useCaseId: env.MAIA_USE_CASE ?? "hub", environment: env.MAIA_ENVIRONMENT ?? "PRD" },
    claudeCode: { baseUrl: env.ANTHROPIC_BASE_URL && !/api\.anthropic\.com/.test(env.ANTHROPIC_BASE_URL) ? env.ANTHROPIC_BASE_URL : "", headers: env.ANTHROPIC_CUSTOM_HEADERS ?? "", keyHelper: env.OKTA_SCRIPT ?? "", key: "" },
  },
  powerbi: { maxRows: 2000, timeoutMs: 20_000, concurrency: 2, cacheMs: 60_000 },
  confluence: { url: env.CONFLUENCE_API_URL ?? "", email: env.CONFLUENCE_EMAIL ?? "", folderId: /^\d+$/.test(folderIdOf(env.CONFLUENCE_FOLDER_ID ?? "")) ? folderIdOf(env.CONFLUENCE_FOLDER_ID!) : "", key: env.CONFLUENCE_API_KEY ?? "" },
});

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
    const saved = await readFile(FILE, "utf8").then(JSON.parse).catch(() => ({}));
    current = merge(defaults(), saved);
  }
  return current!;
}

async function save() {
  await mkdir(DATA_DIR, { recursive: true });
  await writeFile(FILE, JSON.stringify(current, null, 2));
  await chmod(FILE, 0o600).catch(() => undefined);
}

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
    confluence: { url: s.confluence.url, email: s.confluence.email, folderId: s.confluence.folderId, keySet: !!s.confluence.key },
  };
}

/** A folder or page id from what was pasted: the id itself, or a link (…/folder/123…, …/pages/123…). */
const folderIdOf = (v: string) => v.match(/\/(?:folder|pages)\/(\d+)/)?.[1] ?? v.trim();

/** Applies a change from Settings; an empty secret leaves the saved one as it is. */
export async function updateSettings(patch: SettingsPatch) {
  const s = await settings();
  const { anthropicKey, maia, claudeCode, ...ai } = patch.ai ?? {};
  const keep = <T extends Record<string, unknown>>(o: T | undefined) => (o ? Object.fromEntries(Object.entries(o).filter(([, v]) => v !== "" && v !== undefined)) : undefined);
  const next = merge(s, {
    general: patch.general,
    ai: { ...ai, ...(anthropicKey ? { anthropicKey } : {}), maia: maia && { ...maia, ...keep({ clientId: maia.clientId, clientSecret: maia.clientSecret }) }, claudeCode: claudeCode && { ...claudeCode, ...keep({ key: claudeCode.key }) } },
    powerbi: patch.powerbi,
    confluence: patch.confluence && { ...patch.confluence, ...(patch.confluence.folderId !== undefined ? { folderId: folderIdOf(patch.confluence.folderId) } : {}), ...keep({ key: patch.confluence.key }) },
  });
  // Secrets given empty are kept.
  if (maia && !maia.clientId) next.ai.maia.clientId = s.ai.maia.clientId;
  if (maia && !maia.clientSecret) next.ai.maia.clientSecret = s.ai.maia.clientSecret;
  if (claudeCode && !claudeCode.key) next.ai.claudeCode.key = s.ai.claudeCode.key;
  if (patch.confluence && !patch.confluence.key) next.confluence.key = s.confluence.key;
  current = next;
  await save();
  return publicSettings();
}
