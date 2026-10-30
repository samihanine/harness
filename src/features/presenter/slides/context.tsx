import { z } from "zod";
import { pptxImage } from "./images";
import { Icon } from "./picture";
import { Frame, M, Rich, SlideMarkdown, deep, defineTemplate, ease, hex, mix, plain, pptxTitle, rise, tint } from "./kit";

/** Context page: a statement on the left, its elements (title, short text, icon) down a timeline that draws itself. */
export const context = defineTemplate({
  id: "context",
  description: "Context of a project or a subject: an optional statement (intro) on the left, 2-6 elements on the right down a timeline, each a title, a short markdown text and a lucide icon.",
  args: z.object({
    title: z.string(),
    intro: z.string().optional().describe("A short statement leading the page"),
    items: z.array(z.object({ title: z.string(), text: z.string().describe("Markdown, short"), icon: z.string().optional() })).min(2).max(6),
  }),
  example: {
  title: "Contexte",
  intro: "Les équipes lisent les rapports Power BI **sans toujours savoir** d'où viennent les chiffres.",
  items: [
    { title: "Des rapports nombreux", text: "Plus de **120 rapports** publiés, peu documentés.", icon: "layers" },
    { title: "Des questions répétées", text: "Les mêmes questions arrivent chaque semaine aux key users.", icon: "messages-square" },
    { title: "Une licence Pro", text: "Pas de capacité Premium : tout passe par les API standard.", icon: "badge-check" },
    { title: "Des données sensibles", text: "Les accès restent ceux de chaque utilisateur.", icon: "shield-check" },
  ],
},
  Component: ({ args }) => (
    <Frame title={args.title}>
      <div className={`grid h-full gap-[4cqw] ${args.intro ? "grid-cols-[36fr_64fr]" : "grid-cols-1 px-[10cqw]"}`}>
        {args.intro && (
          <M.div variants={rise} className="relative flex items-center">
            <span className="absolute top-[15%] bottom-[15%] left-0 w-[0.4cqw] rounded-full" style={{ background: "var(--s-primary)" }} />
            <p className="pl-[2.4cqw] text-[2.3cqw] leading-[1.25] font-semibold" style={{ fontFamily: "var(--s-title)" }}>
              <Rich text={args.intro} />
            </p>
          </M.div>
        )}
        <div className="relative flex min-h-0 flex-col justify-center gap-[2cqw]">
          <M.span className="absolute top-[2cqw] bottom-[2cqw] left-[1.95cqw] w-[0.16cqw] origin-top" style={{ background: tint(55) }} initial={{ scaleY: 0 }} animate={{ scaleY: 1 }} transition={{ duration: 0.4 + args.items.length * 0.35, ease: "easeInOut", delay: 0.3 }} />
          {args.items.map((it, i) => (
            <M.div key={i} initial={{ opacity: 0, x: "1.5cqw" }} animate={{ opacity: 1, x: 0 }} transition={{ delay: 0.5 + i * 0.35, duration: 0.6, ease }} className="relative flex items-start gap-[1.6cqw]">
              <M.span initial={{ scale: 0 }} animate={{ scale: 1 }} transition={{ delay: 0.45 + i * 0.35, type: "spring", stiffness: 300, damping: 16 }} className="grid size-[4cqw] shrink-0 place-items-center rounded-full border-[0.2cqw] shadow-[0_0.2cqw_0.6cqw_rgba(0,0,0,0.1)]" style={{ borderColor: "var(--s-primary)", background: "var(--s-bg)", color: deep }}>
                <Icon name={it.icon} style={{ width: "1.8cqw", height: "1.8cqw" }} />
              </M.span>
              <div className="flex min-w-0 flex-col gap-[0.3cqw] pt-[0.4cqw]">
                <span className="text-[1.7cqw] leading-tight font-bold">{it.title}</span>
                <div className="text-[1.35cqw] leading-snug opacity-75">
                  <SlideMarkdown text={it.text} />
                </div>
              </div>
            </M.div>
          ))}
        </div>
      </div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const { theme } = ctx;
    const text = hex(theme.text);
    if (args.intro) {
      slide.addShape("rect", { x: 0.8, y: 2.6, w: 0.06, h: 3.4, fill: { color: hex(theme.primary) }, line: { type: "none" } as never });
      slide.addText(plain(args.intro), { x: 1.1, y: 2.2, w: 3.6, h: 4.2, fontSize: 22, bold: true, color: text, fontFace: theme.fontTitle, valign: "middle" });
    }
    const x0 = args.intro ? 5.3 : 2.1;
    const step = Math.min(1.15, 4.6 / args.items.length);
    const top = 2.0 + (4.8 - step * args.items.length) / 2;
    slide.addShape("line", { x: x0 + 0.3, y: top + 0.3, w: 0, h: step * (args.items.length - 1), line: { color: mix(theme.primary, theme.background, 0.55), width: 1.5 } });
    args.items.forEach((it, i) => {
      const y = top + i * step;
      slide.addShape("ellipse", { x: x0, y, w: 0.6, h: 0.6, fill: { color: hex(theme.background) }, line: { color: hex(theme.primary), width: 1.5 } });
      const icon = pptxImage(`icon:${it.icon?.trim() || "sparkles"}`, { x: x0 + 0.15, y: y + 0.15, w: 0.3, h: 0.3 });
      if (icon) slide.addImage(icon);
      slide.addText(
        [
          { text: it.title, options: { bold: true, fontSize: 15, breakLine: true } },
          { text: plain(it.text), options: { fontSize: 12, color: mix(theme.text, theme.background, 0.75) } },
        ],
        { x: x0 + 0.85, y: y - 0.05, w: 12.5 - x0 - 0.85, h: step - 0.05, valign: "top", color: text, fontFace: theme.fontText },
      );
    });
  },
});
