/**
 * The app's options. The values below are the defaults; an admin can change them from Settings ›
 * Config, saved in a JSON file on SharePoint (and cached in this browser), without a developer. They are
 * read at render time: `applyConfig` changes them and the screens are drawn again.
 */

/** Name shown in the header and the browser tab. */
export let APP_TITLE = "Hub";
/** Primary color (buttons, active states), any CSS color. */
export let PRIMARY_COLOR = "#18181b";

/** Top bar (app name, screens, account, settings). Hidden, it still appears when the user must sign in. */
export let SHOW_HEADER = true;

/** AI panel on every screen: false hides every "AI" tab (each screen's own AI flag below also applies). */
export let AI_PANEL = true;

/** Export features and their buttons. */
export const EXPORT = {
  /** Viewer, Guides tab: "PDF" of the clicked visual alone (browser print dialog). */
  visualPdf: true,
  /** Viewer, Guides tab: "PNG" of the clicked visual alone (captured from the tab: the browser asks to share it once). */
  visualPng: false,
  /** Viewer and builder: "Export" button of the report (PDF, .pbix). */
  report: true,
};

/**
 * Screens and panel tabs shown in the app: set one to false to hide it
 * (e.g. VIEWER_TABS.guide = false hides the Guides tab of the viewer). A hidden screen is also closed
 * to direct links (it opens the first screen shown instead). Settings can never be hidden.
 */
export const SCREENS = {
  library: true,
  viewer: true,
  builder: true,
  datasets: true,
  dax: true,
  pivot: true,
};

export const VIEWER_TABS = {
  info: true,
  guide: true,
  ai: true,
};

export const BUILDER_TABS = {
  ai: true,
  inspiration: true,
};

export const DATASET_TABS = {
  ai: true,
  schema: true,
};

export const DAX_TABS = {
  ai: true,
};

export const PIVOT_TABS = {
  fields: true,
  ai: true,
};

/** Viewer, Guides tab: "Explain" button of the clicked visual (the AI reads its data). */
export let VISUAL_EXPLAIN = true;

/** Viewer, Guides tab: live preview of the clicked visual (a second, light embed of that visual alone). */
export let VISUAL_PREVIEW = true;

/** Viewer: guides and info links can be added / edited / deleted (false = read only). */
export let REPORT_VIEWER_EDITOR = true;

/** Instructions added to the AI's own (company vocabulary, tone, rules): for every screen, and per screen. */
export const PROMPTS = { all: "", viewer: "", builder: "", datasets: "", dax: "", pivot: "" };
export type PromptScreen = Exclude<keyof typeof PROMPTS, "all">;
/** A screen's AI instructions with the admin's additions. */
export const withPrompts = (screen: PromptScreen, instructions: string) => {
  const added = [PROMPTS.all, PROMPTS[screen]].filter((s) => s.trim()).join("\n");
  return added ? `${instructions}\n\nAdditional instructions from the administrator (follow them):\n${added}` : instructions;
};

/** Address of each screen. */
export const SCREEN_PATHS: Record<keyof typeof SCREENS, string> = { library: "/", viewer: "/viewer", builder: "/builder", datasets: "/datasets", dax: "/dax", pivot: "/pivot" };
/** First screen shown (where a hidden screen leads). */
export const firstScreen = () => SCREEN_PATHS[(Object.keys(SCREENS) as (keyof typeof SCREENS)[]).find((k) => SCREENS[k]) ?? "library"];

/** Tabs of a panel kept by their flag (`id` = key of the flags); "ai" tabs also need AI_PANEL. */
export const visible = <T extends { id: string }>(tabs: T[], flags: Record<string, boolean>) => tabs.filter((t) => flags[t.id] !== false && (t.id !== "ai" || AI_PANEL));

/* ------------------------------ Runtime config ----------------------------- */

/** Every option, as saved in the JSON file. */
export type HubConfig = {
  appTitle: string;
  primaryColor: string;
  showHeader: boolean;
  aiPanel: boolean;
  visualExplain: boolean;
  visualPreview: boolean;
  viewerEditor: boolean;
  screens: typeof SCREENS;
  viewerTabs: typeof VIEWER_TABS;
  builderTabs: typeof BUILDER_TABS;
  datasetTabs: typeof DATASET_TABS;
  daxTabs: typeof DAX_TABS;
  pivotTabs: typeof PIVOT_TABS;
  export: typeof EXPORT;
  prompts: typeof PROMPTS;
};

export const currentConfig = (): HubConfig =>
  structuredClone({
    appTitle: APP_TITLE,
    primaryColor: PRIMARY_COLOR,
    showHeader: SHOW_HEADER,
    aiPanel: AI_PANEL,
    visualExplain: VISUAL_EXPLAIN,
    visualPreview: VISUAL_PREVIEW,
    viewerEditor: REPORT_VIEWER_EDITOR,
    screens: SCREENS,
    viewerTabs: VIEWER_TABS,
    builderTabs: BUILDER_TABS,
    datasetTabs: DATASET_TABS,
    daxTabs: DAX_TABS,
    pivotTabs: PIVOT_TABS,
    export: EXPORT,
    prompts: PROMPTS,
  });

/** The values of this file, before any change from Settings. */
export const DEFAULT_CONFIG = currentConfig();

const listeners = new Set<() => void>();
/** Called after each change (the root layout draws the screens again). */
export const onConfigChange = (listener: () => void) => (listeners.add(listener), () => void listeners.delete(listener));

/** Only the known options, of the right type, are taken (a hand-edited file cannot break the app). */
const take = <T extends Record<string, unknown>>(target: T, given: unknown) => {
  if (!given || typeof given !== "object") return;
  for (const [k, v] of Object.entries(given)) if (k in target && typeof v === typeof target[k]) (target as Record<string, unknown>)[k] = v;
};

export function applyConfig(given: Partial<HubConfig>, { cache = true } = {}) {
  const pick = <T,>(v: unknown, current: T): T => (typeof v === typeof current ? (v as T) : current);
  APP_TITLE = pick(given.appTitle, APP_TITLE);
  PRIMARY_COLOR = pick(given.primaryColor, PRIMARY_COLOR);
  SHOW_HEADER = pick(given.showHeader, SHOW_HEADER);
  AI_PANEL = pick(given.aiPanel, AI_PANEL);
  VISUAL_EXPLAIN = pick(given.visualExplain, VISUAL_EXPLAIN);
  VISUAL_PREVIEW = pick(given.visualPreview, VISUAL_PREVIEW);
  REPORT_VIEWER_EDITOR = pick(given.viewerEditor, REPORT_VIEWER_EDITOR);
  take(SCREENS, given.screens);
  take(VIEWER_TABS, given.viewerTabs);
  take(BUILDER_TABS, given.builderTabs);
  take(DATASET_TABS, given.datasetTabs);
  take(DAX_TABS, given.daxTabs);
  take(PIVOT_TABS, given.pivotTabs);
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
