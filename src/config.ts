/** Name shown in the header and the browser tab. */
export const APP_TITLE = "Hub";
/** Primary color (buttons, active states), any CSS color. */
export const PRIMARY_COLOR = "#18181b";

/** Top bar (app name, screens, account, AI settings). Hidden, it still appears when the user must sign in. */
export const SHOW_HEADER = true;

/** AI panel on every screen: false hides every "AI" tab (each screen's own AI flag below also applies). */
export const AI_PANEL = true;

/** Export features and their buttons. */
export const EXPORT = {
  /** Viewer, Guides tab: "PDF" of the clicked visual alone (browser print dialog). */
  visualPdf: false,
  /** Viewer, Guides tab: "PNG" of the clicked visual alone (captured from the tab: the browser asks to share it once). */
  visualPng: true,
  /** Viewer and builder: "Export" button of the report (PDF, .pbix). */
  report: true,
};

/**
 * Screens and panel tabs shown in the app: set one to false to hide it
 * (e.g. VIEWER_TABS.guide = false hides the Guides tab of the viewer). A hidden screen is also closed
 * to direct links (it opens the first screen shown instead).
 */
export const SCREENS = {
  library: true,
  viewer: true,
  builder: true,
  datasets: true,
  dax: true,
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

/** Viewer: guides and info links can be added / edited / deleted (false = read only). */
export const REPORT_VIEWER_EDITOR = true;

/** Address of each screen. */
export const SCREEN_PATHS: Record<keyof typeof SCREENS, string> = { library: "/", viewer: "/viewer", builder: "/builder", datasets: "/datasets", dax: "/dax" };
/** First screen shown (where a hidden screen leads). */
export const firstScreen = () => SCREEN_PATHS[(Object.keys(SCREENS) as (keyof typeof SCREENS)[]).find((k) => SCREENS[k]) ?? "library"];

/** Tabs of a panel kept by their flag (`id` = key of the flags); "ai" tabs also need AI_PANEL. */
export const visible = <T extends { id: string }>(tabs: T[], flags: Record<string, boolean>) => tabs.filter((t) => flags[t.id] !== false && (t.id !== "ai" || AI_PANEL));
