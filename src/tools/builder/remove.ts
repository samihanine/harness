/** Removes a visual, or a whole page, from the report being built. */
import { z } from "zod";
import { pageOf, pagesOf, readJson, removePath, visualOf, writeJson } from "@server/pbi/draft";
import { tool } from "../tool";
import { draft } from "./shared";

export const removeItem = tool({
  name: "remove",
  description: "Remove a visual (page + visual) or a whole page (page only) from the report being built.",
  args: z.object({ page: z.string(), visual: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const page = pageOf(d, a.page);
    if (a.visual) {
      const { json } = visualOf(d, page.name, a.visual);
      removePath(d, `definition/pages/${page.name}/visuals/${json.name}`);
      return `Visual ${json.name} removed.`;
    }
    if (pagesOf(d).length === 1) throw new Error("A report keeps at least one page.");
    removePath(d, `definition/pages/${page.name}`);
    const meta = readJson(d, "definition/pages/pages.json");
    meta.pageOrder = (meta.pageOrder ?? []).filter((n: string) => n !== page.name);
    if (meta.activePageName === page.name) meta.activePageName = meta.pageOrder[0];
    writeJson(d, "definition/pages/pages.json", meta);
    return `Page "${page.displayName}" removed.`;
  },
});
