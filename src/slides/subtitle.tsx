import { z } from "zod";
import { H, M, Rich, defineTemplate, ease, hex, pad, plain, rise, stagger } from "./kit";
import { ToileSpray } from "./toile";

/**
 * Subpart title page: its title (and part above it), the trail of the part's subparts — a drop of ink
 * runs to the current one, which fills and keeps a slow halo — and a rose spray growing from the end
 * of the trail (ToileSpray starts at its right end: 58cqw, 49.4cqw). Adapted from the Hub
 * presentation's subsection.
 */
export const subtitle = defineTemplate({
  id: "subtitle",
  description: "Subpart title page (inside a part started by a section page): title, short description, the trail of the part's subparts with the current one lit, and an engraved rose growing from it.",
  kind: "subsection",
  bare: true,
  args: z.object({ title: z.string(), description: z.string().optional().describe("One or two short sentences (**bold**, ==highlight==)") }),
  Component: ({ args, deck }) => {
    const steps = deck.subsections.length ? deck.subsections : [args.title];
    const current = Math.max(0, deck.subsection);
    const at = (i: number) => (steps.length > 1 ? (i / (steps.length - 1)) * 100 : 100);
    const arrive = 2;
    return (
      <div className="relative size-full overflow-hidden">
        <ToileSpray className="pointer-events-none absolute inset-0 size-full" />
        <M.div variants={stagger(0.12, 0.2)} className="absolute top-[3cqw] bottom-[14cqw] left-[12.6cqw] flex w-[48cqw] flex-col justify-center">
          {deck.sections[deck.section] && (
            <M.span variants={rise} className="text-[1cqw] font-semibold tracking-[0.3em] uppercase" style={{ color: "var(--s-primary)" }}>
              {deck.sections[deck.section]}
            </M.span>
          )}
          <M.div variants={{ hidden: { scaleX: 0 }, show: { scaleX: 1, transition: { duration: 0.9, ease } } }} className="mt-[1.2cqw] h-[0.2cqw] w-[5cqw] origin-left rounded-full" style={{ background: "var(--s-primary)" }} />
          <M.h2 variants={rise} className="mt-[2cqw] text-[4.6cqw] leading-[1.04] tracking-[-0.01em] text-balance" style={{ fontFamily: "var(--s-title)" }}>
            <Rich text={args.title} />
          </M.h2>
          {args.description && (
            <M.p variants={rise} className="mt-[2cqw] max-w-[44cqw] text-[1.3cqw] leading-[1.6] opacity-75">
              <Rich text={args.description} />
            </M.p>
          )}
        </M.div>
        <M.div variants={rise} className="absolute bottom-[5.5cqw] left-[8cqw] w-[56cqw]">
          <div className="relative mx-[6cqw] mt-[2.4cqw] h-[2.8cqw]">
            <span className="absolute inset-x-0 top-1/2 h-[0.18cqw] -translate-y-1/2" style={{ background: "var(--s-primary)" }} />
            {current > 0 && (
              <M.span
                className="absolute top-1/2 size-[0.9cqw] -translate-x-1/2 -translate-y-1/2 rounded-full"
                style={{ background: "var(--s-primary)", boxShadow: "0 0 0 0.3cqw var(--s-bg)" }}
                initial={{ left: `${at(current - 1)}%`, opacity: 0 }}
                animate={{ left: `${at(current)}%`, opacity: [0, 1, 1, 0] }}
                transition={{ duration: 1.4, ease, delay: 0.6, opacity: { duration: 1.4, delay: 0.6, times: [0, 0.15, 0.85, 1] } }}
              />
            )}
            {steps.map((step, i) => {
              const done = i < current;
              const here = i === current;
              return (
                <div key={i} className="absolute top-0 -translate-x-1/2" style={{ left: `${at(i)}%` }}>
                  <M.div className="relative size-[2.8cqw]" initial={false} animate={{ scale: here ? 1.3 : 1 }} transition={{ duration: 0.9, ease, delay: here ? arrive - 0.2 : 0 }}>
                    {here && (
                      <M.span
                        className="absolute inset-0 rounded-full border"
                        style={{ borderColor: "var(--s-primary)" }}
                        initial={{ scale: 1, opacity: 0 }}
                        animate={{ scale: [1, 1.9], opacity: [0.55, 0] }}
                        transition={{ duration: 2.6, ease: "easeOut", repeat: Infinity, repeatDelay: 0.6, delay: arrive }}
                      />
                    )}
                    <span
                      className="absolute inset-0 grid place-items-center rounded-full border text-[1.05cqw] transition-colors duration-700"
                      style={{
                        fontFamily: "var(--s-title)",
                        borderColor: here || done ? "var(--s-primary)" : "color-mix(in srgb, var(--s-primary) 35%, transparent)",
                        background: here ? "var(--s-primary)" : done ? "color-mix(in srgb, var(--s-primary) 10%, var(--s-bg))" : "var(--s-bg)",
                        color: here ? "var(--s-bg)" : done ? "var(--s-primary)" : "color-mix(in srgb, var(--s-text) 50%, transparent)",
                      }}
                    >
                      {pad(i + 1)}
                    </span>
                  </M.div>
                  <span
                    className="absolute top-[4.2cqw] left-1/2 w-[13cqw] -translate-x-1/2 text-center text-[0.78cqw] leading-snug tracking-[0.2em] uppercase"
                    style={{ fontWeight: here ? 600 : 400, color: here ? "var(--s-text)" : done ? "color-mix(in srgb, var(--s-primary) 70%, transparent)" : "color-mix(in srgb, var(--s-text) 45%, transparent)" }}
                  >
                    {step}
                  </span>
                </div>
              );
            })}
          </div>
        </M.div>
      </div>
    );
  },
  pptx: (slide, args, { theme, deck }) => {
    if (deck.sections[deck.section]) slide.addText(deck.sections[deck.section].toUpperCase(), { x: 1.6, y: 1.6, w: 7, h: 0.4, fontSize: 12, bold: true, charSpacing: 4, color: hex(theme.primary), fontFace: theme.fontText });
    slide.addText(plain(args.title), { x: 1.6, y: 2.1, w: 7.5, h: 1.4, fontSize: 40, fontFace: theme.fontTitle, color: hex(theme.text), valign: "top" });
    if (args.description) slide.addText(plain(args.description), { x: 1.6, y: 3.6, w: 7, h: 1, fontSize: 16, fontFace: theme.fontText, color: hex(theme.text), transparency: 25, valign: "top" });
    // The trail: a line and a numbered dot per subpart (the current one filled).
    const steps = deck.subsections.length ? deck.subsections : [args.title];
    const x0 = 1.8;
    const x1 = 7.6;
    const y = H - 1.5;
    slide.addShape("line", { x: x0, y, w: x1 - x0, h: 0, line: { color: hex(theme.primary), width: 1.5 } });
    steps.forEach((step, i) => {
      const x = steps.length > 1 ? x0 + ((x1 - x0) * i) / (steps.length - 1) : x1;
      const here = i === Math.max(0, deck.subsection);
      slide.addShape("ellipse", { x: x - 0.18, y: y - 0.18, w: 0.36, h: 0.36, fill: { color: here ? hex(theme.primary) : hex(theme.background) }, line: { color: hex(theme.primary), width: 1 } });
      slide.addText(pad(i + 1), { x: x - 0.18, y: y - 0.18, w: 0.36, h: 0.36, fontSize: 9, align: "center", valign: "middle", color: here ? hex(theme.background) : hex(theme.primary), fontFace: theme.fontText });
      slide.addText(step.toUpperCase(), { x: x - 0.9, y: y + 0.3, w: 1.8, h: 0.4, fontSize: 7, align: "center", charSpacing: 2, color: hex(theme.text), transparency: here ? 0 : 50, fontFace: theme.fontText });
    });
  },
});
