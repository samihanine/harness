import { z } from "zod";
import { Frame, M, Rich, defineTemplate, hex, plain, pptxTitle, rise, stagger } from "./kit";
import { Picture, pictureArgs, pptxPicture } from "./picture";

const column = z.object({
  title: z.string(),
  tone: z.enum(["primary", "muted"]).default("primary").describe("primary: in the accent color (e.g. what comes first); muted: grey (secondary)"),
  items: z.array(z.object({ name: z.string(), text: z.string(), ...pictureArgs })).min(1).max(6),
});

/** Two lists side by side (e.g. priority / secondary), each row a picture (logo, icon), a name and a short text. */
export const lists = defineTemplate({
  id: "lists",
  description: "Two lists side by side, each with a heading and a tone (primary / muted): rows of a picture (logo URL, img:<id>, lucide icon or mascot), a name and a short text — e.g. priority needs on the left, secondary ones on the right.",
  args: z.object({ title: z.string(), left: column, right: column }),
  Component: ({ args }) => (
    <Frame title={args.title}>
      <div className="grid h-full grid-cols-2 gap-[2.4cqw]">
        {[args.left, args.right].map((col, k) => {
          const accent = col.tone === "muted" ? "color-mix(in srgb, var(--s-text) 45%, var(--s-bg))" : "var(--s-primary)";
          const dense = col.items.length > 4;
          return (
            <M.div key={k} variants={stagger(0.08, 0.2 + k * 0.3)} className="flex min-h-0 flex-col overflow-hidden rounded-[1cqw] border" style={{ borderColor: "color-mix(in srgb, var(--s-text) 12%, transparent)" }}>
              <M.div variants={rise} className="px-[1.6cqw] py-[0.9cqw] text-[1.45cqw] font-bold text-white" style={{ background: accent, fontFamily: "var(--s-title)" }}>
                {col.title}
              </M.div>
              {col.items.map((it, i) => (
                <M.div key={i} variants={rise} className={`flex items-center gap-[1.2cqw] border-t px-[1.6cqw] ${dense ? "py-[0.6cqw]" : "py-[1cqw]"}`} style={{ borderColor: "color-mix(in srgb, var(--s-text) 8%, transparent)" }}>
                  <span className="grid size-[3.2cqw] shrink-0 place-items-center">
                    <Picture image={it.image} icon={it.icon} mascot={it.mascot} className={it.icon ? "size-[2.4cqw]" : "max-h-full max-w-full"} />
                  </span>
                  <span className="flex min-w-0 flex-col leading-snug">
                    <span className="text-[1.3cqw] font-semibold"><Rich text={it.name} /></span>
                    <span className="text-[1.05cqw] opacity-70"><Rich text={it.text} /></span>
                  </span>
                </M.div>
              ))}
            </M.div>
          );
        })}
      </div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    [args.left, args.right].forEach((col, k) => {
      const x = 0.8 + k * 6.0;
      const color = col.tone === "muted" ? "94A3B8" : hex(ctx.theme.primary);
      slide.addShape("rect", { x, y: 2.1, w: 5.7, h: 0.5, fill: { color }, line: { color } });
      slide.addText(col.title, { x: x + 0.2, y: 2.1, w: 5.3, h: 0.5, fontSize: 15, bold: true, color: "FFFFFF", fontFace: ctx.theme.fontTitle });
      const h = Math.min(0.8, 4.3 / col.items.length);
      col.items.forEach((it, i) => {
        const y = 2.75 + i * h;
        const image = pptxPicture(it, { x: x + 0.1, y: y + 0.05, w: h - 0.2, h: h - 0.2 });
        if (image) slide.addImage(image);
        slide.addText([{ text: plain(it.name), options: { bold: true, fontSize: 13, breakLine: true } }, { text: plain(it.text), options: { fontSize: 10 } }], { x: x + h, y, w: 5.7 - h, h, valign: "middle", fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
      });
    });
  },
});
