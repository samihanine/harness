/**
 * The Teams assistant: one chat (group or one-to-one) of the signed-in user, watched every few seconds with one
 * light call (its last message); new messages are read and given to the assistant with the conversation before
 * them. It answers when they concern it (its projects, its data, an earlier question, its name) — or says
 * nothing (NO_REPLY). Its answers are sent as the user, signed with its name, short (a chat), with data tables
 * and the mascot's GIFs. It never answers its own messages (their ids are kept), but may answer the user's.
 * Settings and what happened: .local/teams.json.
 */
import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import type { ChatEvent, DatasetRef, ReportRef } from "@shared/types";
import { toStorage } from "../confluence";
import { DATA_DIR } from "../settings";
import { library } from "../library";
import { runChat } from "../agent/run";
import { chatName, messageText, teamsApi, type TeamsMessage } from "./api";
import { MOODS } from "./moods";

export type TeamsConfig = {
  enabled: boolean;
  chatId?: string;
  chatName?: string;
  name: string;
  description: string;
  instructions: string;
  projects: string[];
  reportIds: string[];
  datasetIds: string[];
  /** Answers at most this many times an hour (beyond: silent). */
  maxPerHour: number;
  /** "relevant": whatever concerns it; "mentions": only when its name is written. */
  mode: "relevant" | "mentions";
};
export type TeamsLog = { at: string; from: string; text: string; reply?: string; gif?: string; skipped?: string; error?: string };
type State = { lastTime?: string; lastPreview?: string; sent: string[]; log: TeamsLog[]; replies: number[]; status?: string };

const FILE = join(DATA_DIR, "teams.json");
const GIFS = join(DATA_DIR, "gifs");
const DEFAULT: TeamsConfig = { enabled: false, name: "Scuba", description: "The team's data assistant: it knows the reports, the models and the projects' rules, and answers questions about the figures.", instructions: "", projects: [], reportIds: [], datasetIds: [], maxPerHour: 20, mode: "relevant" };

let config: TeamsConfig = DEFAULT;
let state: State = { sent: [], log: [], replies: [] };
let loaded = false;
let busy = false;

async function load() {
  if (loaded) return;
  loaded = true;
  const f = Bun.file(FILE);
  if (await f.exists()) {
    const saved = await f.json().catch(() => ({}));
    config = { ...DEFAULT, ...(saved.config ?? {}) };
    state = { sent: [], log: [], replies: [], ...(saved.state ?? {}) };
  }
}
async function persist() {
  await mkdir(DATA_DIR, { recursive: true });
  await Bun.write(FILE, JSON.stringify({ config, state: { ...state, sent: state.sent.slice(-300), log: state.log.slice(-60) } }, null, 1));
}

export async function teamsStatus() {
  await load();
  const gifs = await Promise.all(MOODS.map(async (m) => [m, await Bun.file(join(GIFS, `${m}.gif`)).exists()] as const));
  return { config, log: state.log.slice(-40).reverse(), status: state.status, gifs: Object.fromEntries(gifs) };
}

export async function updateTeams(patch: Partial<TeamsConfig>) {
  await load();
  const switched = patch.chatId !== undefined && patch.chatId !== config.chatId;
  config = { ...config, ...patch };
  // A new chat (or switched on): only the messages from now on.
  if (switched || (patch.enabled && !state.lastTime)) Object.assign(state, { lastTime: new Date().toISOString(), lastPreview: undefined });
  state.status = config.enabled ? (config.chatId ? "watching" : "choose a chat") : "off";
  await persist();
  return teamsStatus();
}

export async function saveGif(mood: string, data: Uint8Array) {
  if (!MOODS.includes(mood)) throw new Error("Unknown mood");
  await mkdir(GIFS, { recursive: true });
  await Bun.write(join(GIFS, `${mood}.gif`), data);
}

export async function teamsChats() {
  const api = teamsApi();
  const me = await api.me();
  return (await api.chats()).map((c) => ({ id: c.id, name: chatName(c, me.displayName), type: c.chatType, last: c.lastMessagePreview?.createdDateTime }));
}

/* ----------------------------------- Loop ----------------------------------- */

const line = (m: TeamsMessage, bot: string) => {
  const t = new Date(m.createdDateTime).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" });
  return `[${t}] ${state.sent.includes(m.id) ? `${bot} (you)` : (m.from?.user?.displayName ?? "?")}: ${messageText(m).replace(/^🤖\s*\S+\s*/, "")}`;
};

/** One look at the chat: nothing new (one light call), or the new messages handled. */
export async function teamsTick() {
  await load();
  if (!config.enabled || !config.chatId || busy) return;
  const api = teamsApi();
  busy = true;
  try {
    const last = await api.lastMessage(config.chatId);
    if (!last || last.id === state.lastPreview) return;
    state.lastPreview = last.id;
    const recent = (await api.messages(config.chatId, 30)).filter((m) => m.messageType === "message" && !m.deletedDateTime).sort((a, b) => a.createdDateTime.localeCompare(b.createdDateTime));
    const fresh = recent.filter((m) => (!state.lastTime || m.createdDateTime > state.lastTime) && !state.sent.includes(m.id) && !messageText(m).startsWith(`🤖 ${config.name}`));
    if (recent.length) state.lastTime = recent.at(-1)!.createdDateTime;
    if (!fresh.length) return void (await persist());
    await respond(fresh, recent);
    await persist();
  } catch (e) {
    state.status = `error: ${e instanceof Error ? e.message : e}`;
  } finally {
    busy = false;
  }
}

async function respond(fresh: TeamsMessage[], recent: TeamsMessage[]) {
  const api = teamsApi();
  const chatId = config.chatId!;
  const text = fresh.map((m) => messageText(m)).join(" ");
  const entry: TeamsLog = { at: new Date().toISOString(), from: [...new Set(fresh.map((m) => m.from?.user?.displayName ?? "?"))].join(", "), text: text.slice(0, 500) };
  state.log.push(entry);
  if (config.mode === "mentions" && !new RegExp(`\\b${config.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i").test(text)) return void (entry.skipped = "not mentioned");
  state.replies = state.replies.filter((t) => Date.now() - t < 3600_000);
  if (state.replies.length >= config.maxPerHour) return void (entry.skipped = `limit of ${config.maxPerHour} answers an hour`);
  // What it works with: the projects, the catalog's reports and models chosen.
  const [reports, models] = await Promise.all([library("report").catch(() => []), library("dataset").catch(() => [])]);
  const references: ReportRef[] = reports.filter((e) => e.id && config.reportIds.includes(e.id)).map((e) => ({ id: e.id!, groupId: e.groupId, appId: e.appId, name: e.name }));
  const datasets: DatasetRef[] = models.filter((e) => e.id && config.datasetIds.includes(e.id)).map((e) => ({ id: e.id!, groupId: e.groupId, name: e.name }));
  const before = recent.filter((m) => !fresh.includes(m)).slice(-15);
  const message = [
    before.length ? `The conversation before (oldest first):\n${before.map((m) => line(m, config.name)).join("\n")}` : "",
    `New message${fresh.length > 1 ? "s" : ""} (ids ${fresh.map((m) => m.id).join(", ")}):\n${fresh.map((m) => line(m, config.name)).join("\n")}`,
    "Answer them now as the chat's assistant, or reply exactly NO_REPLY when nothing calls for you.",
  ]
    .filter(Boolean)
    .join("\n\n");
  let reply = "";
  let gif: string | undefined;
  const tables = new Map<number, string>();
  const emit = (e: ChatEvent) => {
    if (e.type === "text") reply += e.text;
    else if (e.type === "effect" && e.name === "gif") gif = String(e.data);
    else if (e.type === "effect" && e.name === "table") tables.set((e.data as { n: number }).n, (e.data as { html: string }).html);
    else if (e.type === "error") entry.error = e.error;
  };
  await runChat({ conversationId: `teams:${chatId}`, screen: "teams", message, scope: { doc: { name: config.name, description: config.description, instructions: config.instructions, chat: config.chatName, chatId }, projects: config.projects, references, datasets } }, emit, new AbortController().signal);
  reply = reply.trim();
  if (!reply || /^NO_REPLY\b/.test(reply)) return void (entry.skipped = "nothing for it");
  // Tables where the reply put them ({{table:1}}), the others after it.
  let html = toStorage(reply.replace(/\{\{table:(\d+)\}\}/g, (_, n) => `TABLE_${n}_HERE`));
  for (const [n, t] of tables) html = html.includes(`TABLE_${n}_HERE`) ? html.replace(new RegExp(`(<p>)?TABLE_${n}_HERE(</p>)?`), t) : `${html}${t}`;
  const signed = `<p>🤖 <b>${config.name}</b></p>${html}`;
  const gifFile = gif && MOODS.includes(gif) ? Bun.file(join(GIFS, `${gif}.gif`)) : undefined;
  const bytes = gifFile && (await gifFile.exists()) ? new Uint8Array(await gifFile.arrayBuffer()) : undefined;
  const id = await api.send(chatId, signed, bytes);
  state.sent.push(id);
  state.replies.push(Date.now());
  Object.assign(entry, { reply: reply.slice(0, 800), gif: bytes ? gif : undefined });
}

let timer: ReturnType<typeof setInterval> | undefined;
/** Watches the chat every 4 seconds while the server runs (when switched on). */
export function startTeams() {
  timer ??= setInterval(() => void teamsTick(), 4000);
  timer.unref?.();
}
