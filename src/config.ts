/** Name shown in the header and the browser tab. */
export const APP_TITLE = "Hub";
/** Primary color (buttons, active states), any CSS color. */
export const PRIMARY_COLOR = "#18181b";

/**
 * Screens and panel tabs shown in the app: set one to false to hide it
 * (e.g. VIEWER_TABS.guide = false hides the Guides tab of the viewer).
 */
export const SCREENS = {
  library: true,
  viewer: true,
  builder: true,
  datasets: true,
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
  model: true,
};

/** Viewer: guides and info links can be added / edited / deleted (false = read only). */
export const REPORT_VIEWER_EDITOR = true;

/** Tabs of a panel kept by their flag (`id` = key of the flags). */
export const visible = <T extends { id: string }>(tabs: T[], flags: Record<string, boolean>) => tabs.filter((t) => flags[t.id] !== false);
