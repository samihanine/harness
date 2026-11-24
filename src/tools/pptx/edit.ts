/** Changes to the uploaded PowerPoint, shape by shape (everything else kept as it is). */
import { z } from "zod";
import { capture, captureArgs } from "../deck/capture-powerbi";
import { tool, type ToolContext } from "../tool";
import { openPptx, pictureBytes, savePptx } from "./shared";

const source = z.object({ capture: z.object(captureArgs).optional().describe("A fresh picture of a Power BI page or visual."), src: z.string().optional().describe("file:<id> (a capture or an uploaded picture) or an https URL.") });

async function sourceBytes(ctx: ToolContext, s: z.infer<typeof source>) {
  if (s.capture) return (await capture(ctx, s.capture)).shot.png;
  if (s.src) return pictureBytes(s.src);
  throw new Error("Give capture or src.");
}

export const setText = tool({
  name: "set_text",
  description: "Replace the text of a shape (or of a table cell: row, col from 0), keeping its formatting (font, size, color of its first run). Lines become paragraphs.",
  args: z.object({ slide: z.number().int().min(1), shape: z.string(), text: z.string(), row: z.number().int().optional(), col: z.number().int().optional() }),
  changes: true,
  async run(a, ctx) {
    const { doc, pptx } = await openPptx(ctx);
    await pptx.setText(a.slide, a.shape, a.text, a.row !== undefined && a.col !== undefined ? { row: a.row, col: a.col } : undefined);
    await savePptx(ctx, doc, pptx);
    return `Text of shape ${a.shape} on slide ${a.slide} replaced.`;
  },
});

export const replacePicture = tool({
  name: "replace_picture",
  description:
    "Replace a picture of a slide, keeping its frame: e.g. an old Power BI screenshot by a fresh capture of the same report page / visual with coherent filters (capture), or by another picture (src). fit cover (default): cropped to the frame; contain: the frame takes the picture's proportions.",
  args: z.object({ slide: z.number().int().min(1), shape: z.string(), ...source.shape, fit: z.enum(["cover", "contain"]).optional() }),
  changes: true,
  async run(a, ctx) {
    const bytes = await sourceBytes(ctx, a);
    const { doc, pptx } = await openPptx(ctx);
    await pptx.replacePicture(a.slide, a.shape, bytes, a.fit ?? "cover");
    await savePptx(ctx, doc, pptx);
    return `Picture ${a.shape} on slide ${a.slide} replaced. Check it with look_at_slides.`;
  },
});

export const addPicture = tool({
  name: "add_picture",
  description: "Add a picture on a slide at a place in inches (height from the picture's proportions when not given).",
  args: z.object({ slide: z.number().int().min(1), ...source.shape, x: z.number(), y: z.number(), w: z.number(), h: z.number().optional() }),
  changes: true,
  async run(a, ctx) {
    const bytes = await sourceBytes(ctx, a);
    const { doc, pptx } = await openPptx(ctx);
    const id = await pptx.addPicture(a.slide, bytes, { x: a.x, y: a.y, w: a.w, h: a.h });
    await savePptx(ctx, doc, pptx);
    return `Picture added on slide ${a.slide}: shape id ${id}.`;
  },
});

export const addText = tool({
  name: "add_text",
  description: "Add a text box on a slide (inches; size in points; color #RRGGBB).",
  args: z.object({ slide: z.number().int().min(1), text: z.string(), x: z.number(), y: z.number(), w: z.number(), h: z.number(), size: z.number().optional(), bold: z.boolean().optional(), color: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const { doc, pptx } = await openPptx(ctx);
    const id = await pptx.addText(a.slide, a.text, { x: a.x, y: a.y, w: a.w, h: a.h }, { size: a.size, bold: a.bold, color: a.color });
    await savePptx(ctx, doc, pptx);
    return `Text box added on slide ${a.slide}: shape id ${id}.`;
  },
});

export const shapeAction = tool({
  name: "shape",
  description: "Move / resize a shape (x, y, w, h in inches; what is not given stays) or delete it.",
  args: z.object({ slide: z.number().int().min(1), shape: z.string(), action: z.enum(["move", "delete"]), x: z.number().optional(), y: z.number().optional(), w: z.number().optional(), h: z.number().optional() }),
  changes: true,
  async run(a, ctx) {
    const { doc, pptx } = await openPptx(ctx);
    if (a.action === "delete") await pptx.deleteShape(a.slide, a.shape);
    else await pptx.moveShape(a.slide, a.shape, a);
    await savePptx(ctx, doc, pptx);
    return `Shape ${a.shape} ${a.action === "delete" ? "deleted" : "moved"}.`;
  },
});

export const slideAction = tool({
  name: "slide",
  description: "Duplicate a slide (the copy comes right after it), delete it, or move it to position `to`.",
  args: z.object({ slide: z.number().int().min(1), action: z.enum(["duplicate", "delete", "move"]), to: z.number().int().optional() }),
  changes: true,
  async run(a, ctx) {
    const { doc, pptx } = await openPptx(ctx);
    let note = "";
    if (a.action === "duplicate") note = `the copy is slide ${await pptx.duplicateSlide(a.slide)}`;
    else if (a.action === "delete") await pptx.deleteSlide(a.slide);
    else await pptx.moveSlide(a.slide, a.to ?? 1);
    await savePptx(ctx, doc, pptx);
    return `Slide ${a.slide}: ${a.action} done${note ? `, ${note}` : ""}. Numbers of the following slides changed: read the structure again.`;
  },
});
