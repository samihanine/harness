import { z } from "zod";
import { Frame, M, Rich, defineTemplate, ease, hex, plain, pptxTitle } from "./kit";

/**
 * A roadmap: numbered steps along a line drawn from left to right, each milestone appearing when the
 * line reaches it, the labels alternately above and below; `current` highlights where things stand.
 */
export const roadmap = defineTemplate({
  id: "roadmap",
  description: "Animated roadmap: 3-8 steps (title + optional short text) along a line drawn from left to right, labels above and below in turn; current: the step in progress (1-based, optional).",
  args: z.object({ title: z.string(), steps: z.array(z.object({ title: z.string(), text: z.string().optional() })).min(3).max(8), current: z.number().optional() }),
  Component: ({ args }) => {
    const n = args.steps.length;
    const step = 2.6 / n;
    return (
      <Frame title={args.title}>
        <div className="relative h-full">
          {/* The road: a thick line drawn across, then the milestones on it. */}
          <M.div className="absolute top-1/2 right-[3%] left-[3%] h-[0.5cqw] origin-left -translate-y-1/2 rounded-full" style={{ background: "linear-gradient(90deg, var(--s-primary), var(--s-secondary))" }} initial={{ scaleX: 0 }} animate={{ scaleX: 1 }} transition={{ duration: 2.6, ease: "easeInOut", delay: 0.3 }} />
          {args.steps.map((s, i) => {
            const x = 3 + (94 * (i + 0.5)) / n;
            const above = i % 2 === 0;
            const delay = 0.3 + step * (i + 0.5);
            const current = args.current === i + 1;
            const done = args.current !== undefined && i + 1 < args.current;
            return (
              <div key={i} className="absolute top-1/2" style={{ left: `${x}%` }}>
                <M.span
                  className="absolute grid size-[4cqw] -translate-x-1/2 -translate-y-1/2 place-items-center rounded-full border-[0.35cqw] text-[1.6cqw] font-bold"
                  style={{ borderColor: "var(--s-primary)", background: current || done ? "var(--s-primary)" : "var(--s-bg)", color: current || done ? "#fff" : "var(--s-primary)", fontFamily: "var(--s-title)" }}
                  initial={{ scale: 0 }}
                  animate={current ? { scale: [0, 1.15, 1, 1.08, 1] } : { scale: 1 }}
                  transition={current ? { duration: 2.2, delay, times: [0, 0.2, 0.3, 0.65, 1], repeat: Infinity, repeatDelay: 1.2 } : { type: "spring", stiffness: 260, damping: 16, delay }}
                >
                  {i + 1}
                </M.span>
                <M.span className="absolute w-[0.25cqw] -translate-x-1/2" style={{ background: "color-mix(in srgb, var(--s-text) 25%, transparent)", height: "3cqw", [above ? "bottom" : "top"]: "2cqw" }} initial={{ scaleY: 0 }} animate={{ scaleY: 1 }} transition={{ duration: 0.4, delay: delay + 0.2 }} />
                <M.div
                  className="absolute flex w-[15cqw] -translate-x-1/2 flex-col gap-[0.3cqw] text-center"
                  style={{ [above ? "bottom" : "top"]: "5.3cqw" }}
                  initial={{ opacity: 0, y: above ? "1cqw" : "-1cqw" }}
                  animate={{ opacity: 1, y: 0 }}
                  transition={{ duration: 0.6, delay: delay + 0.35, ease }}
                >
                  <span className="text-[1.55cqw] leading-tight font-bold" style={{ color: current ? "var(--s-primary)" : undefined }}>
                    <Rich text={s.title} />
                  </span>
                  {s.text && <span className="text-[1.15cqw] leading-snug opacity-70"><Rich text={s.text} /></span>}
                </M.div>
              </div>
            );
          })}
        </div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const n = args.steps.length;
    const y = 4.4;
    slide.addShape("rect", { x: 0.8, y: y - 0.04, w: 11.7, h: 0.08, fill: { color: hex(ctx.theme.primary) }, line: { type: "none" } as never });
    args.steps.forEach((s, i) => {
      const x = 0.8 + (11.7 * (i + 0.5)) / n;
      const above = i % 2 === 0;
      slide.addShape("ellipse", { x: x - 0.28, y: y - 0.28, w: 0.56, h: 0.56, fill: { color: args.current === i + 1 ? hex(ctx.theme.primary) : "FFFFFF" }, line: { color: hex(ctx.theme.primary), width: 2 } });
      slide.addText(String(i + 1), { x: x - 0.28, y: y - 0.28, w: 0.56, h: 0.56, align: "center", valign: "middle", fontSize: 14, bold: true, color: args.current === i + 1 ? "FFFFFF" : hex(ctx.theme.primary) });
      slide.addText([{ text: plain(s.title), options: { bold: true, fontSize: 12, breakLine: true } }, ...(s.text ? [{ text: plain(s.text), options: { fontSize: 9 } }] : [])], { x: x - 0.9, y: above ? y - 1.65 : y + 0.45, w: 1.8, h: 1.2, align: "center", valign: above ? "bottom" : "top", fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    });
  },
});
