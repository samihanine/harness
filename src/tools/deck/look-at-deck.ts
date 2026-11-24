/** Pictures of the deck as PowerPoint draws it (the exported .pptx, converted by Microsoft). */
import { z } from "zod";
import { deckToPptx } from "@server/pptx/deck";
import { slidePictures } from "@server/pptx/render";
import { base64, tool } from "../tool";
import { deckOf } from "./deck";

export const lookAtDeck = tool({
  name: "look_at_deck",
  description: "See slides of the deck as pictures, exactly as the PowerPoint export draws them (up to 6 per call). Check layout, overflowing text, pictures, after each change.",
  args: z.object({ slides: z.array(z.number().int().min(1)).max(6).optional().describe("Slide numbers (default: the first 6).") }),
  async run(a, ctx) {
    ctx.status("Drawing the slides…");
    const pptx = await deckToPptx(deckOf(ctx));
    const wanted = a.slides?.length ? a.slides : [1, 2, 3, 4, 5, 6];
    const pics = await slidePictures(new Uint8Array(pptx), wanted, 1000);
    return { text: `Slides ${wanted.slice(0, pics.length).join(", ")}:`, images: pics.map((p) => ({ mediaType: "image/png", data: base64(p) })) };
  },
});
