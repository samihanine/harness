/** Places the visuals of a page in rows: margins, even gaps, nothing overlapping or outside the page. */
import { z } from "zod";
import { pageOf, visualOf, visualsOf, writeJson } from "@server/pbi/draft";
import { layoutProblems, layoutRows, layoutRowsIn, snapBox } from "@server/pbi/layout";
import { tool } from "../tool";
import { draft } from "./shared";

export const layoutPage = tool({
  name: "layout_page",
  description:
    "Lay out a page in rows, top to bottom: each row has a height in px (omit to share the rest) and its visuals left to right with relative widths (1 = equal share). E.g. a row of 4 cards (height 110), then a chart (width 2) beside a table (width 1). Every visual of the page should be in a row — or give \"area\" (x, y, width, height) to lay out only these visuals inside it (e.g. cards inside a frame shape, under a header placed with place_visuals).",
  args: z.object({
    page: z.string(),
    rows: z.array(z.object({ height: z.number().optional(), items: z.array(z.object({ visual: z.string(), width: z.number().optional() })) })),
    margin: z.number().optional(),
    gap: z.number().optional(),
    area: z.object({ x: z.number(), y: z.number(), width: z.number(), height: z.number() }).optional(),
  }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const page = pageOf(d, a.page);
    const size = { width: page.width ?? 1280, height: page.height ?? 720 };
    const boxes = a.area ? layoutRowsIn(a.rows, a.area, { margin: a.margin ?? 12, gap: a.gap ?? 12 }) : layoutRows(a.rows, size, { margin: a.margin ?? 20, gap: a.gap ?? 16 });
    for (const b of boxes) {
      const { path, json } = visualOf(d, page.name, b.visual);
      json.position = { ...json.position, x: b.x, y: b.y, width: b.width, height: b.height };
      writeJson(d, path, json);
    }
    return `Placed ${boxes.length} visuals on "${page.displayName}".`;
  },
});

export const placeVisuals = tool({
  name: "place_visuals",
  description:
    "Free positions on a page, in px: each visual's x, y, width, height (snapped to a 4 px grid, kept in the page), and its layer: \"back\" (under the others: a frame, a band, a background picture) or \"front\". Visuals may lie over a shape, an image or a text box drawn under them (layers); other overlaps are reported. For a header, a frame with cards inside, a logo in a corner. Rows of visuals: layout_page.",
  args: z.object({
    page: z.string(),
    visuals: z.array(z.object({ visual: z.string(), x: z.number(), y: z.number(), width: z.number(), height: z.number(), layer: z.enum(["back", "front"]).optional() })),
  }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const page = pageOf(d, a.page);
    const size = { width: page.width ?? 1280, height: page.height ?? 720 };
    const all = visualsOf(d, page.name);
    const zs = all.map((v) => v.json.position?.z ?? 0);
    let back = Math.min(0, ...zs) - 1000;
    let front = Math.max(0, ...zs) + 1000;
    for (const b of a.visuals) {
      const { path, json } = visualOf(d, page.name, b.visual);
      const box = snapBox(b, size);
      const z = b.layer === "back" ? (back -= 10) : b.layer === "front" ? (front += 10) : json.position?.z ?? 0;
      json.position = { ...json.position, x: box.x, y: box.y, width: box.width, height: box.height, z, tabOrder: json.position?.tabOrder ?? z };
      writeJson(d, path, json);
    }
    const problems = layoutProblems(visualsOf(d, page.name).filter((v) => v.json.position && !v.json.parentGroupName).map((v) => ({ name: v.json.name, ...v.json.position, hidden: v.json.isHidden, type: v.json.visual?.visualType })), size);
    return `Placed ${a.visuals.length} visual(s) on "${page.displayName}".${problems.length ? ` To fix before saving: ${problems.join("; ")}.` : " No overlap."}`;
  },
});
