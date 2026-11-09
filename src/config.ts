import { useSyncExternalStore } from "react";

/**
 * The app's options. The values below are the defaults; an admin can change them from Settings
 * (saved in a JSON file on SharePoint, and cached in this browser), without a developer. They are
 * read at render time: `applyConfig` changes them and the screens are drawn again.
 */

/** Name shown in the sidebar and the browser tab. */
export let APP_TITLE = "Hub";
/** Primary color (buttons, active states), any CSS color. */
export let PRIMARY_COLOR = "#2563eb";

/** The local token server (scripts/token-server.ts): Microsoft access tokens for the browser. */
export let TOKEN_SERVER = "http://localhost:3301";

/** AI panel: files added to a conversation (the paperclip). */
export let FILE_UPLOAD = true;

/**
 * Limits that keep the semantic models fast for everyone (docs/performance.md): rows a query returns,
 * time before a query is stopped, queries at once per model, the pause between two, how long an
 * identical query's result is reused, and Power BI reports embedded at once (slides).
 */
export const PERFORMANCE = { maxRows: 500, timeoutMs: 10_000, concurrency: 4, intervalMs: 50, cacheMs: 600_000, maxEmbeds: 4 };

/** Home page ("/"): a card per screen; false opens the first screen shown instead. */
export let HOME = true;

/** AI panel on every screen: false hides every AI chat (the screens' own AI flags below also apply). */
export let AI_PANEL = true;

/** Report viewer › Guides: live preview of the clicked visual (a second, light embed of that visual alone). */
export let VISUAL_PREVIEW = true;
/** Report viewer › Guides: "Explain" button of the clicked visual (the AI reads its data). */
export let VISUAL_EXPLAIN = true;
/** Report viewer › Guides: guides written in the viewer (title, content, new guide rows, ids); false = read only. */
export let GUIDE_EDITOR = true;
/** AI: the failed tool calls are kept in this browser and the latest ones shown to the AI as mistakes not to repeat (Settings downloads them). */
export let LEARN_FROM_MISTAKES = true;
/** AI panel: the mascot shows what the assistant does (thinking, working, talking, sleeping), dressed for each screen. */
export let MASCOT = true;
/** AI panel: thumbs up / down (with an optional note) on each answer and on the whole conversation. */
export let RATINGS = true;
/** Report viewer › Guides: "Why did it change?" button of a time chart (the app breaks the change down by dimension, the AI tells it). */
export let VISUAL_WHY = true;

/** Export features and their buttons. */
export const EXPORT = {
  /** Report viewer › Guides: "PDF" of the clicked visual alone (browser print dialog). */
  visualPdf: true,
  /** Report viewer › Guides: "PNG" of the clicked visual alone (captured from the tab: the browser asks to share it once). */
  visualPng: false,
  /** Report viewer and editor: "Export" button of the report (PDF, .pbix). */
  report: true,
  /** Data Scientist: "Excel" button of a pivot card (a workbook with the pivot, connected to the model). */
  pivotExcel: true,
};

/**
 * Screens shown in the app: set one to false to hide it. A hidden screen is also closed to direct
 * links (it opens the first screen shown instead). Settings can never be hidden.
 */
export const SCREENS = {
  viewer: true,
  editor: true,
  research: true,
  slides: true,
  reports: true,
  datasets: true,
  excel: true,
  /** Confluence pages (Settings › Confluence), edited as markdown; each one is a context linked to items. */
  confluence: true,
};
export type Screen = keyof typeof SCREENS;

/** Tabs of the Builder (the report itself is always there): its report measures, the reference reports given to the AI. */
export const BUILDER_TABS = {
  measures: true,
  reports: true,
};

/** Tabs of the report viewer (the report itself is always there). */
export const VIEWER_TABS = {
  guides: true,
};

/** Instructions added to the AI's own (company vocabulary, tone, rules): for every screen, and per screen. */
export const PROMPTS = { all: "", viewer: "", editor: "", research: "", slides: "", reports: "", datasets: "", excel: "", confluence: "" };
export type PromptScreen = Exclude<keyof typeof PROMPTS, "all">;
/** A screen's AI instructions with the admin's additions. */
export const withPrompts = (screen: PromptScreen, instructions: string) => {
  const added = [PROMPTS.all, PROMPTS[screen]].filter((s) => s.trim()).join("\n");
  return added ? `${instructions}\n\nAdditional instructions from the administrator (follow them):\n${added}` : instructions;
};

/* ------------------------------ Runtime config ----------------------------- */

/** Every option, as saved in the JSON file. */
export type HubConfig = {
  appTitle: string;
  primaryColor: string;
  tokenServer: string;
  fileUpload: boolean;
  performance: typeof PERFORMANCE;
  home: boolean;
  aiPanel: boolean;
  visualPreview: boolean;
  visualExplain: boolean;
  guideEditor: boolean;
  learnFromMistakes: boolean;
  visualWhy: boolean;
  ratings: boolean;
  mascot: boolean;
  screens: typeof SCREENS;
  viewerTabs: typeof VIEWER_TABS;
  builderTabs: typeof BUILDER_TABS;
  export: typeof EXPORT;
  prompts: typeof PROMPTS;
};

export const currentConfig = (): HubConfig =>
  structuredClone({
    appTitle: APP_TITLE,
    primaryColor: PRIMARY_COLOR,
    tokenServer: TOKEN_SERVER,
    fileUpload: FILE_UPLOAD,
    performance: PERFORMANCE,
    home: HOME,
    aiPanel: AI_PANEL,
    visualPreview: VISUAL_PREVIEW,
    visualExplain: VISUAL_EXPLAIN,
    guideEditor: GUIDE_EDITOR,
    learnFromMistakes: LEARN_FROM_MISTAKES,
    visualWhy: VISUAL_WHY,
    ratings: RATINGS,
    mascot: MASCOT,
    screens: SCREENS,
    viewerTabs: VIEWER_TABS,
    builderTabs: BUILDER_TABS,
    export: EXPORT,
    prompts: PROMPTS,
  });

/** The values of this file, before any change from Settings. */
export const DEFAULT_CONFIG = currentConfig();

const listeners = new Set<() => void>();
let version = 0;
/** Called after each change. */
export const onConfigChange = (listener: () => void) => (listeners.add(listener), () => void listeners.delete(listener));
/**
 * Draws the calling component again when an option changes (the layout, each screen's host and the
 * Workspace use it): the screens are redrawn in place, never mounted again, so embedded reports and
 * chats keep their state.
 */
export const useConfig = () => useSyncExternalStore(onConfigChange, () => version);

/** Only the known options, of the right type, are taken (a hand-edited file cannot break the app). */
const take = <T extends Record<string, unknown>>(target: T, given: unknown) => {
  if (!given || typeof given !== "object") return;
  for (const [k, v] of Object.entries(given)) if (k in target && typeof v === typeof target[k]) (target as Record<string, unknown>)[k] = v;
};

export function applyConfig(given: Partial<HubConfig>, { cache = true } = {}) {
  const pick = <T,>(v: unknown, current: T): T => (typeof v === typeof current ? (v as T) : current);
  APP_TITLE = pick(given.appTitle, APP_TITLE);
  PRIMARY_COLOR = pick(given.primaryColor, PRIMARY_COLOR);
  TOKEN_SERVER = pick(given.tokenServer, TOKEN_SERVER);
  FILE_UPLOAD = pick(given.fileUpload, FILE_UPLOAD);
  take(PERFORMANCE, given.performance);
  HOME = pick(given.home, HOME);
  AI_PANEL = pick(given.aiPanel, AI_PANEL);
  VISUAL_PREVIEW = pick(given.visualPreview, VISUAL_PREVIEW);
  VISUAL_EXPLAIN = pick(given.visualExplain, VISUAL_EXPLAIN);
  GUIDE_EDITOR = pick(given.guideEditor, GUIDE_EDITOR);
  LEARN_FROM_MISTAKES = pick(given.learnFromMistakes, LEARN_FROM_MISTAKES);
  VISUAL_WHY = pick(given.visualWhy, VISUAL_WHY);
  RATINGS = pick(given.ratings, RATINGS);
  MASCOT = pick(given.mascot, MASCOT);
  take(SCREENS, given.screens);
  take(VIEWER_TABS, given.viewerTabs);
  take(BUILDER_TABS, given.builderTabs);
  take(EXPORT, given.export);
  take(PROMPTS, given.prompts);
  if (typeof document !== "undefined") {
    document.title = APP_TITLE;
    document.body?.style.setProperty("--color-primary", PRIMARY_COLOR);
  }
  if (cache) {
    // Only what differs from config.ts: a later change of a default in the code still applies.
    try {
      const diff = changes(currentConfig(), DEFAULT_CONFIG);
      if (Object.keys(diff).length) localStorage.setItem(CACHE, JSON.stringify(diff));
      else localStorage.removeItem(CACHE);
    } catch {}
  }
  version++;
  listeners.forEach((l) => l());
}

/** The options of `next` that differ from `base` (nested objects compared key by key). */
function changes(next: Record<string, unknown>, base: Record<string, unknown>) {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(next)) {
    if (v && typeof v === "object") {
      const inner = changes(v as Record<string, unknown>, (base[k] ?? {}) as Record<string, unknown>);
      if (Object.keys(inner).length) out[k] = inner;
    } else if (v !== base[k]) out[k] = v;
  }
  return out;
}

/** Last config seen in this browser: applied at start, before the SharePoint file is read again. */
const CACHE = "hub-config";
try {
  const cached = typeof localStorage !== "undefined" ? localStorage.getItem(CACHE) : null;
  if (cached) applyConfig(JSON.parse(cached), { cache: false });
} catch {}
