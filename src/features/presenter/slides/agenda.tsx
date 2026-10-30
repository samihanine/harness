import { z } from "zod";
import { pptxImage } from "./images";
import { H, M, W, brandOf, defineTemplate, ease, hex, plain, reveal, rise, stagger } from "./kit";

/**
 * Agenda page: the brand and a fine rule on top, the word "Agenda" written vertically, the parts as a
 * numbered list (from the deck's part pages, or given) with the current one lit on a band, and a
 * picture on the right. Naming a part, it starts that part (the deck's outline comes from them).
 */
export const agenda = defineTemplate({
  id: "agenda",
  description:
    "Agenda page, also the page starting each part: the parts numbered (filled automatically, or given in items), the current one highlighted, the word in the title written vertically, a picture on the right (URL). With \"part\" (the part's name) it STARTS that part: put one before each part; without, it is the deck's summary.",
  kind: "summary",
  kindOf: (args) => (args.part ? "section" : "summary"),
  bare: true,
  args: z.object({
    title: z.string().default("Agenda").describe("Written vertically on the left"),
    part: z.string().optional().describe("The part this page starts (it is then highlighted); none: the deck's summary"),
    image: z.string().optional().describe("Picture URL (right part)"),
    items: z.array(z.string()).max(8).optional().describe("Entries (default: the deck's parts)"),
    current: z.number().optional().describe("Entry to highlight, from 1 (default: the next part)"),
  }),
  Component: ({ args, theme, deck }) => {
    const entries = args.items?.length ? args.items : deck.outline.map((p) => p.title).filter(Boolean);
    const current = (args.current ?? (args.part ? deck.section : deck.next) + 1) - 1;
    return (
      <div className={`grid size-full ${args.image ? "grid-cols-[63fr_37fr]" : "grid-cols-1"}`}>
        <div className="relative">
          <M.div variants={stagger(0.1, 0.1)} className="absolute top-[3.6cqw] right-[4cqw] left-[3.2cqw] flex items-end gap-[2.4cqw]">
            <M.span variants={rise} className="shrink-0 text-[4.2cqw] leading-none" style={{ fontFamily: "var(--s-brand)" }}>
              {brandOf(theme)}
            </M.span>
            <M.div variants={{ hidden: { scaleX: 0 }, show: { scaleX: 1, transition: { duration: 1.2, ease } } }} className="mb-[0.4cqw] h-px flex-1 origin-left" style={{ background: "color-mix(in srgb, var(--s-text) 40%, transparent)" }} />
          </M.div>
          {/* Written bottom to top, along the list, faint (as a margin title). */}
          <M.span
            variants={{ hidden: { opacity: 0, y: "2cqw" }, show: { opacity: 0.32, y: 0, transition: { duration: 1, ease, delay: 0.3 } } }}
            className="absolute top-[30%] bottom-[18%] left-[2.4cqw] flex items-center text-[4.6cqw] leading-none tracking-wide whitespace-nowrap"
            style={{ fontFamily: "var(--s-title)" }}
          >
            {/* The animation owns the outer transform: the rotation sits on the text itself. */}
            <span style={{ writingMode: "vertical-rl", transform: "rotate(180deg)" }}>{args.title}</span>
          </M.span>
          <M.ol variants={stagger(0.12, 0.5)} className="absolute top-1/2 right-0 left-[11cqw] flex -translate-y-[30%] flex-col gap-[1.6cqw]">
            {entries.map((e, i) => {
              const here = i === current;
              return (
                <M.li key={i} variants={rise} className="flex items-center gap-[1.4cqw] py-[0.6cqw] pr-[2cqw]" style={{ background: here ? "color-mix(in srgb, var(--s-text) 7%, transparent)" : undefined }}>
                  <span className="grid size-[3.4cqw] shrink-0 place-items-center text-[1.6cqw] text-white" style={{ background: here ? "var(--s-secondary)" : "color-mix(in srgb, var(--s-text) 38%, var(--s-bg))" }}>
                    {i + 1}
                  </span>
                  <span className="text-[1.9cqw] leading-tight" style={{ fontWeight: here ? 700 : 400 }}>
                    {e}
                  </span>
                </M.li>
              );
            })}
            {!entries.length && <M.li variants={rise} className="text-[1.4cqw] opacity-50">Add agenda pages naming their part: the agenda lists them.</M.li>}
          </M.ol>
        </div>
        {args.image && (
          <M.div variants={reveal} className="overflow-hidden">
            <M.img src={args.image} alt="" className="size-full object-cover" variants={{ hidden: { scale: 1.12 }, show: { scale: 1, transition: { duration: 2.4, ease } } }} />
          </M.div>
        )}
      </div>
    );
  },
  pptx: (slide, args, { theme, deck }) => {
    const entries = args.items?.length ? args.items : deck.outline.map((p) => p.title).filter(Boolean);
    const current = (args.current ?? (args.part ? deck.section : deck.next) + 1) - 1;
    const w = args.image ? W * 0.63 : W;
    const image = args.image && pptxImage(args.image, { x: w, y: 0, w: W - w, h: H }, "cover");
    if (image) slide.addImage(image);
    slide.addText(brandOf(theme), { x: 0.4, y: 0.35, w: 3, h: 0.8, fontSize: 32, fontFace: "Bodoni MT", color: hex(theme.text) });
    slide.addShape("line", { x: 2.6, y: 1.05, w: w - 3.1, h: 0, line: { color: hex(theme.text), width: 0.5, transparency: 50 } });
    slide.addText(args.title, { x: -1.6, y: 4.2, w: 4, h: 0.8, fontSize: 34, rotate: 270, color: hex(theme.text), transparency: 65, fontFace: theme.fontTitle });
    entries.forEach((e, i) => {
      const y = 2.9 + i * 0.75;
      if (i === current) slide.addShape("rect", { x: 1.5, y: y - 0.08, w: w - 1.7, h: 0.62, fill: { color: hex(theme.text), transparency: 93 }, line: { type: "none" } as never });
      slide.addShape("rect", { x: 1.6, y, w: 0.46, h: 0.46, fill: { color: i === current ? hex(theme.secondary ?? "#8b5a2b") : "9CA3AF" }, line: { type: "none" } as never });
      slide.addText(String(i + 1), { x: 1.6, y, w: 0.46, h: 0.46, fontSize: 14, color: "FFFFFF", align: "center", valign: "middle" });
      slide.addText(plain(e), { x: 2.25, y, w: w - 2.6, h: 0.46, fontSize: 17, bold: i === current, valign: "middle", fontFace: theme.fontText, color: hex(theme.text) });
    });
  },
});
