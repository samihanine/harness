/**
 * The app's options. The values below are the defaults; an admin can change them from Settings
 * (saved in a JSON file on SharePoint, and cached in this browser), without a developer. They are
 * read at render time: `applyConfig` changes them and the screens are drawn again.
 */

/** Name shown in the sidebar and the browser tab. */
export let APP_TITLE = "Hub";
/** Primary color (buttons, active states), any CSS color. */
export let PRIMARY_COLOR = "#2563eb";

/** AI panel on every screen: false hides every AI chat (the screens' own AI flags below also apply). */
export let AI_PANEL = true;

/** Report viewer › Guides: live preview of the clicked visual (a second, light embed of that visual alone). */
export let VISUAL_PREVIEW = true;
/** Report viewer › Guides: "Explain" button of the clicked visual (the AI reads its data). */
export let VISUAL_EXPLAIN = true;
/** Report viewer › Guides: guides written in the viewer (title, content, new guide rows, ids); false = read only. */
export let GUIDE_EDITOR = true;
/** Excel editor: uploading files into image / file fields (to the schema's uploads folder, set in the Schema editor). */
export let EXCEL_UPLOAD = true;

/** Export features and their buttons. */
export const EXPORT = {
  /** Report viewer › Guides: "PDF" of the clicked visual alone (browser print dialog). */
  visualPdf: true,
  /** Report viewer › Guides: "PNG" of the clicked visual alone (captured from the tab: the browser asks to share it once). */
  visualPng: false,
  /** Report viewer and editor: "Export" button of the report (PDF, .pbix). */
  report: true,
};

/**
 * Screens shown in the app: set one to false to hide it. A hidden screen is also closed to direct
 * links (it opens the first screen shown instead). Settings can never be hidden.
 */
export const SCREENS = {
  viewer: true,
  editor: true,
  research: true,
  reports: true,
  datasets: true,
  excel: true,
  contexts: true,
  slides: true,
};
export type Screen = keyof typeof SCREENS;

/** Tabs of the report viewer (the report itself is always there). */
export const VIEWER_TABS = {
  guides: true,
};

/** Instructions added to the AI's own (company vocabulary, tone, rules): for every screen, and per screen. */
export const PROMPTS = { all: "", viewer: "", editor: "", research: "", reports: "", datasets: "", excel: "", contexts: "", slides: "" };
export type PromptScreen = Exclude<keyof typeof PROMPTS, "all">;
/** A screen's AI instructions with the admin's additions. */
export const withPrompts = (screen: PromptScreen, instructions: string) => {
  const added = [PROMPTS.all, PROMPTS[screen]].filter((s) => s.trim()).join("\n");
  return added ? `${instructions}\n\nAdditional instructions from the administrator (follow them):\n${added}` : instructions;
};

/** Address and name of each screen. */
export const SCREEN_PATHS: Record<Screen, string> = { viewer: "/viewer", editor: "/editor", research: "/research", reports: "/reports", datasets: "/datasets", excel: "/excel", contexts: "/contexts", slides: "/slides" };
export const SCREEN_LABELS: Record<Screen, string> = { viewer: "Viewer", editor: "Builder", research: "Research", reports: "Reports", datasets: "Datasets", excel: "Excel", contexts: "Contexts", slides: "Slides" };
/** First screen shown (where a hidden screen leads). */
export const firstScreen = () => SCREEN_PATHS[(Object.keys(SCREENS) as Screen[]).find((k) => SCREENS[k]) ?? "viewer"] ?? "/settings";

/* ------------------------------ Runtime config ----------------------------- */

/** Every option, as saved in the JSON file. */
export type HubConfig = {
  appTitle: string;
  primaryColor: string;
  aiPanel: boolean;
  visualPreview: boolean;
  visualExplain: boolean;
  guideEditor: boolean;
  excelUpload: boolean;
  screens: typeof SCREENS;
  viewerTabs: typeof VIEWER_TABS;
  export: typeof EXPORT;
  prompts: typeof PROMPTS;
};

export const currentConfig = (): HubConfig =>
  structuredClone({
    appTitle: APP_TITLE,
    primaryColor: PRIMARY_COLOR,
    aiPanel: AI_PANEL,
    visualPreview: VISUAL_PREVIEW,
    visualExplain: VISUAL_EXPLAIN,
    guideEditor: GUIDE_EDITOR,
    excelUpload: EXCEL_UPLOAD,
    screens: SCREENS,
    viewerTabs: VIEWER_TABS,
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
  AI_PANEL = pick(given.aiPanel, AI_PANEL);
  VISUAL_PREVIEW = pick(given.visualPreview, VISUAL_PREVIEW);
  VISUAL_EXPLAIN = pick(given.visualExplain, VISUAL_EXPLAIN);
  GUIDE_EDITOR = pick(given.guideEditor, GUIDE_EDITOR);
  EXCEL_UPLOAD = pick(given.excelUpload, EXCEL_UPLOAD);
  take(SCREENS, given.screens);
  take(VIEWER_TABS, given.viewerTabs);
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
