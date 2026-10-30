import { z } from "zod";
import { H, M, Rich, W, brandOf, defineTemplate, ease, hex, plain, rise, stagger } from "./kit";
import { Monogram } from "./monogram";
import { ToileFlowers } from "./toile";

/** A hairline of the frame, drawn from a corner (`axis`: which way it grows). */
const line = (axis: "x" | "y", delay: number) => ({
  hidden: axis === "x" ? { scaleX: 0 } : { scaleY: 0 },
  show: { ...(axis === "x" ? { scaleX: 1 } : { scaleY: 1 }), transition: { duration: 2.4, ease, delay } },
});

/**
 * Title page (the cover, or a closing page): toile de Jouy flowers all around the title, a fine
 * double frame drawn from the corners, the brand letter. Adapted from the Hub presentation's cover.
 */
export const title = defineTemplate({
  id: "title",
  description: "Title page with engraved flowers all around the title (a cover, or a closing \"Questions?\" page): title, optional eyebrow, subtitle, top-right context and footer (author · date).",
  kind: "cover",
  bare: true,
  args: z.object({ title: z.string(), subtitle: z.string().optional(), eyebrow: z.string().optional(), context: z.string().optional().describe("Small text top right, e.g. \"Reporting · Q3\""), footer: z.string().optional().describe("e.g. author · date") }),
  Component: ({ args, theme, deck }) => (
    <div className="relative size-full overflow-hidden">
      <ToileFlowers className="pointer-events-none absolute inset-0 size-full" />
      {[
        ["2.2cqw", 0.55],
        ["2.65cqw", 0.25],
      ].map(([inset, opacity], i) => (
        <M.div key={i} className="pointer-events-none absolute" style={{ inset: inset as string }} variants={{ hidden: { opacity: 0 }, show: { opacity: opacity as number, transition: { duration: 1.2, delay: 0.3 + i * 0.3 } } }}>
          <M.span className="absolute inset-x-0 top-0 h-px origin-left" style={{ background: "var(--s-primary)" }} variants={line("x", 0.4 + i * 0.3)} />
          <M.span className="absolute inset-x-0 bottom-0 h-px origin-right" style={{ background: "var(--s-primary)" }} variants={line("x", 0.4 + i * 0.3)} />
          <M.span className="absolute inset-y-0 left-0 w-px origin-bottom" style={{ background: "var(--s-primary)" }} variants={line("y", 0.4 + i * 0.3)} />
          <M.span className="absolute inset-y-0 right-0 w-px origin-top" style={{ background: "var(--s-primary)" }} variants={line("y", 0.4 + i * 0.3)} />
        </M.div>
      ))}
      <M.div variants={rise} className="absolute top-[4.4cqw] left-[5.4cqw]">
        <Monogram letter={brandOf(theme).charAt(0)} className="h-[3.4cqw] w-auto" />
      </M.div>
      {args.context && (
        <M.p variants={rise} className="absolute top-[5cqw] right-[5.4cqw] text-[0.8cqw] tracking-[0.34em] uppercase opacity-60">
          {args.context}
        </M.p>
      )}
      <M.div variants={stagger(0.15, 0.4)} className="absolute top-[49%] left-1/2 flex w-[52cqw] -translate-x-1/2 -translate-y-1/2 flex-col items-center text-center">
        {args.eyebrow && (
          <M.span variants={rise} className="text-[1cqw] font-semibold tracking-[0.3em] uppercase" style={{ color: "var(--s-primary)" }}>
            {args.eyebrow}
          </M.span>
        )}
        <M.div variants={{ hidden: { scaleX: 0 }, show: { scaleX: 1, transition: { duration: 0.9, ease } } }} className="mt-[1.2cqw] h-[0.2cqw] w-[5cqw] rounded-full" style={{ background: "var(--s-primary)" }} />
        <M.h1 variants={rise} className="mt-[2cqw] text-[5.2cqw] leading-[1.02] font-normal tracking-[-0.01em] text-balance" style={{ fontFamily: "var(--s-title)" }}>
          <Rich text={args.title} />
        </M.h1>
        {args.subtitle && (
          <M.p variants={rise} className="mt-[2cqw] max-w-[38cqw] text-[1.5cqw] leading-snug italic opacity-70">
            <Rich text={args.subtitle} />
          </M.p>
        )}
      </M.div>
      {args.footer && (
        <M.p variants={rise} className="absolute inset-x-[5.4cqw] bottom-[4.4cqw] text-[0.78cqw] tracking-[0.3em] uppercase opacity-60">
          {args.footer}
        </M.p>
      )}
    </div>
  ),
  // PowerPoint: the flowers as a picture (drawn before the export), the frame and the centered title.
  pptx: (slide, args, { theme, deck, toile }) => {
    if (toile) slide.addImage({ data: toile, x: 0, y: 0, w: W, h: H });
    for (const [inset, transparency] of [[0.3, 45], [0.36, 75]] as const)
      slide.addShape("rect", { x: inset, y: inset, w: W - 2 * inset, h: H - 2 * inset, fill: { type: "none" } as never, line: { color: hex(theme.primary), width: 0.75, transparency } });
    slide.addText(brandOf(theme).charAt(0), { x: 0.7, y: 0.55, w: 0.8, h: 0.6, fontSize: 30, fontFace: "Bodoni MT", color: hex(theme.primary) });
    if (args.context) slide.addText(args.context.toUpperCase(), { x: W - 4.7, y: 0.6, w: 4, h: 0.3, fontSize: 9, align: "right", charSpacing: 4, color: hex(theme.text), transparency: 40, fontFace: theme.fontText });
    if (args.eyebrow) slide.addText(args.eyebrow.toUpperCase(), { x: 2, y: 2.3, w: W - 4, h: 0.4, fontSize: 12, bold: true, align: "center", charSpacing: 4, color: hex(theme.primary), fontFace: theme.fontText });
    slide.addText(plain(args.title), { x: 1.5, y: 2.8, w: W - 3, h: 1.6, fontSize: 48, align: "center", fontFace: theme.fontTitle, color: hex(theme.text) });
    if (args.subtitle) slide.addText(plain(args.subtitle), { x: 2.5, y: 4.4, w: W - 5, h: 0.8, fontSize: 18, italic: true, align: "center", fontFace: theme.fontText, color: hex(theme.text), transparency: 30 });
    if (args.footer) slide.addText(args.footer.toUpperCase(), { x: 0.7, y: H - 0.9, w: W - 1.4, h: 0.3, fontSize: 9, charSpacing: 4, color: hex(theme.text), transparency: 40, fontFace: theme.fontText });
  },
});
