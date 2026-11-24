/** Formatting of a visual (title, labels, colors, legend…) from plain values, checked against Power BI's format catalog. */
import { z } from "zod";
import { applyAdds, pageOf, visualOf, writeJson } from "@server/pbi/draft";
import { formatOps, formatOptions as options, ignoredText, pageFormatOps } from "@server/pbi/visual-spec";
import { tool } from "../tool";
import { draft } from "./shared";

export const formatOptions = tool({
  name: "format_options",
  description: "The format objects of a visual type (title, labels, legend, dataPoint…), or the properties of some objects with their type (give object: \"title, labels\").",
  args: z.object({ type: z.string(), object: z.string().optional() }),
  async run(a) {
    return options(a.type, a.object);
  },
});

export const formatVisual = tool({
  name: "format_visual",
  description:
    'Format a visual: {"title": {"show": true, "text": "Sales", "fontSize": 14, "fontColor": "#203247"}, "labels": {"show": true}, "dataPoint": {"defaultColor": "#B39251"}, "background": {"color": "#FFFFFF"}}. Names come from format_options. Also page colors: give page_background / page_wallpaper with no visual.',
  args: z.object({ page: z.string(), visual: z.string().optional(), format: z.record(z.string(), z.record(z.string(), z.unknown())).optional(), page_background: z.string().optional(), page_wallpaper: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const page = pageOf(d, a.page);
    if (!a.visual) {
      const path = `definition/pages/${page.name}/page.json`;
      const json = JSON.parse(JSON.stringify(d.parts.find((p) => p.path === path) ? (await import("@server/pbi/draft")).readJson(d, path) : {}));
      applyAdds(json, pageFormatOps(json, { background: a.page_background, wallpaper: a.page_wallpaper }));
      writeJson(d, path, json);
      return `Page "${page.displayName}" colors set.`;
    }
    const { path, json } = visualOf(d, page.name, a.visual);
    const { ops, ignored } = formatOps(json, a.format ?? {});
    applyAdds(json, ops);
    writeJson(d, path, json);
    return `Formatted ${json.name} (${ops.length} properties).${ignoredText(ignored)}`;
  },
});
