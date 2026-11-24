/** The user's view of the embedded report (from the browser) as text for the AI. */
export type ViewState = {
  pages: { name: string; displayName: string; active: boolean; hidden?: boolean }[];
  reportFilters: string[];
  pageFilters: string[];
  slicers: { name: string; title?: string; fields: { table: string; column: string }[]; selected: string; hidden?: boolean }[];
  visuals: { name: string; type: string; title?: string }[];
};

export const viewText = (v: ViewState) =>
  [
    `Pages: ${v.pages.map((p) => `"${p.displayName}" (${p.name}${p.active ? ", CURRENT" : ""}${p.hidden ? ", hidden" : ""})`).join(", ")}`,
    `Slicers on the current page:${v.slicers.length ? "" : " none"}`,
    ...v.slicers.map((s) => `- "${s.title ?? ""}" name=${s.name} on ${s.fields.map((f) => `'${f.table}'[${f.column}]`).join(" > ") || "?"}: ${s.selected}${s.hidden ? " (hidden)" : ""}`),
    `Report filters: ${v.reportFilters.join("; ") || "none"}`,
    `Page filters: ${v.pageFilters.join("; ") || "none"}`,
    `Visuals: ${v.visuals.map((x) => `${x.type} "${x.title ?? ""}" (${x.name})`).join(", ")}`,
  ].join("\n");
