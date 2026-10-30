import { useState } from "react";
import { z } from "zod";
import { Frame, M, defineTemplate, fade, hex, pptxTitle, stagger } from "./kit";
import { PowerBiFrame, cropSchema, pptxPowerBi } from "./powerbi";

const side = z.object({ label: z.string(), report: z.string().describe("Report id or name"), page: z.string().optional(), crop: cropSchema.optional() });

/**
 * Two Power BI views laid one over the other with a slider between them (before / after, two pages or
 * visuals): dragging the line uncovers the left view or the right one.
 */
export const compare = defineTemplate({
  id: "compare",
  description: "Two Power BI views laid one over the other (before/after, two pages or visuals of the same size), each with a label, and a slider the presenter drags to compare them.",
  args: z.object({ title: z.string(), left: side, right: side }),
  Component: ({ args }) => <Slider {...args} />,
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    [args.left, args.right].forEach((s, i) => {
      const x = 0.8 + i * 6.0;
      slide.addText(s.label.toUpperCase(), { x, y: 2.1, w: 5.7, h: 0.4, fontSize: 12, bold: true, color: hex(i ? ctx.theme.primary : ctx.theme.text), fontFace: ctx.theme.fontText });
      pptxPowerBi(slide, ctx.theme, s, { x, y: 2.6, w: 5.7, h: 4.2 });
    });
  },
});

/** The two views, the right one cut at the slider (dragged with the mouse or a finger). */
function Slider({ title, left, right }: { title: string; left: z.infer<typeof side>; right: z.infer<typeof side> }) {
  const [at, setAt] = useState(0.5);
  const drag = (e: React.PointerEvent<HTMLDivElement>) => {
    const box = (e.currentTarget.parentElement as HTMLElement).getBoundingClientRect();
    setAt(Math.min(1, Math.max(0, (e.clientX - box.left) / box.width)));
  };
  return (
    <Frame title={title}>
      <M.div variants={stagger(0.2, 0.3)} className="flex h-full flex-col gap-[0.8cqw]">
        <M.div variants={fade} className="flex justify-between text-[1.2cqw] font-semibold tracking-wide uppercase">
          <span style={{ opacity: 0.6 }}>{left.label}</span>
          <span style={{ color: "var(--s-primary)" }}>{right.label}</span>
        </M.div>
        <M.div variants={fade} className="relative min-h-0 flex-1">
          <div className="absolute inset-0">
            <PowerBiFrame report={left.report} page={left.page} crop={left.crop} />
          </div>
          <div className="absolute inset-0">
            <PowerBiFrame
              report={right.report}
              page={right.page}
              crop={right.crop}
              reveal={at}
              overlay={() => (
                <div
                  className="absolute top-0 bottom-0 flex w-[28px] -translate-x-1/2 cursor-ew-resize touch-none justify-center"
                  style={{ left: `${at * 100}%`, pointerEvents: "auto" }}
                  onPointerDown={(e) => (e.currentTarget.setPointerCapture(e.pointerId), drag(e))}
                  onPointerMove={(e) => e.currentTarget.hasPointerCapture(e.pointerId) && drag(e)}
                >
                  <div className="h-full w-[3px] bg-white shadow-[0_0_6px_rgba(0,0,0,0.45)]" />
                  <div className="absolute top-1/2 grid size-[28px] -translate-y-1/2 place-items-center rounded-full bg-white text-[13px] font-bold text-slate-700 shadow-[0_1px_6px_rgba(0,0,0,0.4)]">⇔</div>
                </div>
              )}
            />
          </div>
        </M.div>
      </M.div>
    </Frame>
  );
}
