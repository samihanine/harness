/** The uploaded PowerPoint's structure, its slides as pictures, and the pictures it holds. */
import { z } from "zod";
import { readStoredFile } from "@server/files";
import { slidePictures } from "@server/pptx/render";
import { base64, tool } from "../tool";
import { openPptx, pptxDoc } from "./shared";

export const pptxStructure = tool({
  name: "pptx_structure",
  description: "The PowerPoint's slides: layout, shapes (id, name, kind text/picture/table/chart, position in inches, text, table cells, picture size), notes. Shape ids are what the edit tools take.",
  args: z.object({ slides: z.array(z.number().int().min(1)).optional() }),
  async run(a, ctx) {
    const { pptx } = await openPptx(ctx);
    const all = await pptx.structure();
    const list = a.slides?.length ? all.filter((s) => a.slides!.includes(s.number)) : all;
    return [
      `${all.length} slides (13.33 × 7.5 in when wide).`,
      ...list.map((s) =>
        [
          `## Slide ${s.number}${s.layout ? ` (layout "${s.layout}")` : ""}`,
          ...s.shapes.map((x) => `- id ${x.id} ${x.kind}${x.placeholder ? ` [${x.placeholder}]` : ""} "${x.name}" at ${x.x},${x.y} size ${x.w}×${x.h}${x.text ? `: ${JSON.stringify(x.text.slice(0, 300))}` : ""}${x.rows ? `: table ${JSON.stringify(x.rows).slice(0, 400)}` : ""}${x.image ? ` picture ${x.image.file} ${x.image.width}×${x.image.height}px` : ""}`),
          s.notes ? `notes: ${s.notes.slice(0, 300)}` : "",
        ].filter(Boolean).join("\n"),
      ),
    ].join("\n");
  },
});

export const lookAtSlides = tool({
  name: "look_at_slides",
  description: "See slides of the PowerPoint as pictures (as PowerPoint draws them), up to 6 per call.",
  args: z.object({ slides: z.array(z.number().int().min(1)).max(6).optional() }),
  async run(a, ctx) {
    const d = pptxDoc(ctx);
    ctx.status("Drawing the slides…");
    const wanted = a.slides?.length ? a.slides : [1, 2, 3, 4, 5, 6];
    const pics = await slidePictures((await readStoredFile(d.fileId)).data, wanted, 1000);
    return { text: `Slides ${wanted.slice(0, pics.length).join(", ")}:`, images: pics.map((p) => ({ mediaType: "image/png", data: base64(p) })) };
  },
});

export const viewPicture = tool({
  name: "view_picture",
  description: "See one picture of a slide in full size (e.g. to recognise which Power BI report and page a screenshot comes from, and its filters).",
  args: z.object({ slide: z.number().int().min(1), shape: z.string() }),
  async run(a, ctx) {
    const { pptx } = await openPptx(ctx);
    const s = (await pptx.structure())[a.slide - 1];
    const shape = s?.shapes.find((x) => x.id === a.shape);
    if (!shape?.image) throw new Error(`Shape ${a.shape} of slide ${a.slide} is not a picture.`);
    const bytes = await pptx.media(shape.image.file);
    const mediaType = bytes[0] === 0x89 ? "image/png" : bytes[0] === 0xff ? "image/jpeg" : "image/gif";
    return { text: `Picture ${shape.image.file} (${shape.image.width}×${shape.image.height}px), shown at ${shape.w}×${shape.h} in.`, images: [{ mediaType, data: base64(bytes) }] };
  },
});
