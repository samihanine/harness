/** Places the visuals of a page in rows: margins, even gaps, nothing overlapping or outside the page. */
import { z } from "zod";
import { pageOf, visualOf, writeJson } from "@server/pbi/draft";
import { layoutRows } from "@server/pbi/layout";
import { tool } from "../tool";
import { draft } from "./shared";

export const layoutPage = tool({
  name: "layout_page",
  description:
    "Lay out a page in rows, top to bottom: each row has a height in px (omit to share the rest) and its visuals left to right with relative widths (1 = equal share). E.g. a row of 4 cards (height 110), then a chart (width 2) beside a table (width 1). Every visual of the page should be in a row.",
  args: z.object({
    page: z.string(),
    rows: z.array(z.object({ height: z.number().optional(), items: z.array(z.object({ visual: z.string(), width: z.number().optional() })) })),
    margin: z.number().optional(),
    gap: z.number().optional(),
  }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const page = pageOf(d, a.page);
    const boxes = layoutRows(a.rows, { width: page.width ?? 1280, height: page.height ?? 720 }, { margin: a.margin ?? 20, gap: a.gap ?? 16 });
    for (const b of boxes) {
      const { path, json } = visualOf(d, page.name, b.visual);
      json.position = { ...json.position, x: b.x, y: b.y, width: b.width, height: b.height };
      writeJson(d, path, json);
    }
    return `Placed ${boxes.length} visuals on "${page.displayName}".`;
  },
});
