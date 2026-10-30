import { z } from "zod";
import { Frame, M, Rich, defineTemplate, hex, plain, pptxTitle, rise, stagger } from "./kit";
import { Picture, pictureArgs, pptxPicture } from "./picture";

/** Cards (1-6: one row up to 4, then rows of 3): a picture (image, icon or mascot), a title and a text each. */
export const cards = defineTemplate({
  id: "cards",
  description: "1-6 cards (one row up to 4, rows of 3 above): each a title, a text and an optional picture on top (image URL / img:<id>, lucide icon, or the mascot in an outfit); optional lead sentence.",
  args: z.object({
    title: z.string(),
    lead: z.string().optional(),
    pictures: z.enum(["logo", "photo"]).default("logo").describe("logo: the picture whole, small (logos, icons); photo: across the top of the card, cropped"),
    items: z.array(z.object({ title: z.string(), text: z.string(), ...pictureArgs })).min(1).max(6),
  }),
  Component: ({ args }) => {
    const columns = args.items.length <= 4 ? args.items.length : 3;
    const pictures = args.items.some((c) => c.image || c.icon || c.mascot);
    return (
    <Frame title={args.title}>
      <div className="flex h-full flex-col gap-[2.4cqw]">
        {args.lead && <M.p variants={rise} className="text-[1.6cqw] leading-snug opacity-80"><Rich text={args.lead} /></M.p>}
        <M.div variants={stagger(0.12, 0.3)} className="grid min-h-0 flex-1 gap-[1.6cqw]" style={{ gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))` }}>
          {args.items.map((c, i) => (
            <M.div key={i} variants={rise} className={`flex min-h-0 flex-col gap-[0.9cqw] rounded-[1.2cqw] ${args.items.length > 4 ? "px-[1.6cqw] pt-[1.4cqw] pb-[2cqw]" : "px-[2cqw] pt-[2cqw] pb-[2.6cqw]"}`} style={{ background: "color-mix(in srgb, var(--s-primary) 7%, var(--s-bg))", borderTop: "0.35cqw solid var(--s-primary)" }}>
              {pictures &&
                (args.pictures === "photo" && c.image ? (
                  <Picture image={c.image} cover className="h-[14cqw] w-full shrink-0 rounded-[0.6cqw]" />
                ) : (
                  <span className={`flex shrink-0 items-center justify-center ${args.items.length > 4 ? "h-[4cqw]" : c.mascot ? "h-[12cqw]" : "h-[8cqw]"}`}>
                    <Picture image={c.image} icon={c.icon} mascot={c.mascot} className={c.mascot ? "h-full w-auto" : c.icon ? "size-[4cqw]" : "h-full max-w-full rounded-[0.6cqw]"} />
                  </span>
                ))}
              <span className={`${args.items.length > 4 ? "text-[1.5cqw]" : "text-[1.7cqw]"} font-bold`} style={{ fontFamily: "var(--s-title)" }}><Rich text={c.title} /></span>
              <span className={`${args.items.length > 4 ? "text-[1.1cqw]" : "text-[1.25cqw]"} leading-snug opacity-75`}><Rich text={c.text} /></span>
            </M.div>
          ))}
        </M.div>
      </div>
    </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const top = args.lead ? 3.0 : 2.3;
    if (args.lead) slide.addText(plain(args.lead), { x: 0.8, y: 2.2, w: 11.5, h: 0.6, fontSize: 16, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    const gap = 0.25;
    const columns = args.items.length <= 4 ? args.items.length : 3;
    const rows = Math.ceil(args.items.length / columns);
    const w = (11.7 - gap * (columns - 1)) / columns;
    const h = (6.8 - top - gap * (rows - 1)) / rows;
    args.items.forEach((c, i) => {
      const x = 0.8 + (i % columns) * (w + gap);
      const y = top + Math.floor(i / columns) * (h + gap);
      slide.addShape("roundRect", { x, y, w, h, fill: { color: hex(ctx.theme.primary), transparency: 92 }, line: { color: hex(ctx.theme.primary), transparency: 100 }, rectRadius: 0.12 });
      slide.addShape("rect", { x, y, w, h: 0.06, fill: { color: hex(ctx.theme.primary) }, line: { color: hex(ctx.theme.primary) } });
      const photo = args.pictures === "photo";
      // A photo across the card, the mascot large, a logo or an icon small.
      const size = photo ? 1.7 : c.mascot ? 1.5 : 0.8;
      const image = photo ? pptxPicture(c, { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: size }, "cover") : pptxPicture(c, { x: x + w / 2 - size / 2, y: y + 0.2, w: size, h: size });
      if (image) slide.addImage(image);
      const textTop = y + (image ? size + 0.35 : 0.25);
      slide.addText([{ text: plain(c.title), options: { bold: true, fontSize: rows > 1 ? 15 : 18, breakLine: true } }, { text: plain(c.text), options: { fontSize: rows > 1 ? 11 : 13 } }], { x: x + 0.2, y: textTop, w: w - 0.4, h: y + h - textTop - 0.15, valign: "top", fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    });
  },
});
