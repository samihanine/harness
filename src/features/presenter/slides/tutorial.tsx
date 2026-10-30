import { z } from "zod";
import { Frame, M, SlideMarkdown, defineTemplate, ease, fitText, hex, plain, pptxTitle, rise, stagger } from "./kit";
import { PowerBiFrame, cropSchema, pptxPowerBi, unionZone, useReportRef, visualZone } from "./powerbi";

const item = z.object({
  text: z.string().describe("Markdown: a **bold** first line as its title, then a short explanation (lists allowed)"),
  type: z.enum(["default", "info"]).default("default").describe("default: numbered in the second accent; info: a tip, in green"),
  visual: z.string().optional().describe("The visual it explains (its id from the report's schema): framed on the page with the item's number"),
  visuals: z.array(z.string()).optional().describe("Several visuals explained together (ids): ONE frame around all of them"),
});

/** Space the frame keeps around its visuals (page px): a little more above and below. */
const PAD = { x: 6, y: 12 };

/**
 * Each item's visuals (a visual already explained by an earlier item is left to it) and its number:
 * one number per framed visual or group, as on a printed guide; items without a visual get a dot.
 * When no item points to a visual, every item is numbered.
 */
function numbering(items: z.infer<typeof item>[]) {
  const seen = new Set<string>();
  const visuals = items.map((it) => [...(it.visuals ?? []), ...(it.visual ? [it.visual] : [])].filter((v) => !seen.has(v) && seen.add(v)));
  const framed = visuals.some((v) => v.length);
  let n = 0;
  return items.map((_, i) => ({ visuals: visuals[i], n: !framed || visuals[i].length ? ++n : undefined }));
}

/** The colors of an item's number and frame. */
const tone = (type: string) => (type === "info" ? "var(--s-success)" : "var(--s-secondary)");

/** A number in a filled circle (the item's badge, on the panel and on its visual's frame). */
function Badge({ n, type, className = "" }: { n?: number; type: string; className?: string }) {
  // No number (an item without its visual): a small dot in its place.
  if (n === undefined) return <span className={`mx-[0.8cqw] mt-[0.55cqw] size-[0.8cqw] shrink-0 rounded-full ${className}`} style={{ background: tone(type) }} />;
  return (
    <span className={`grid shrink-0 place-items-center rounded-full font-semibold text-white shadow-[0_0.2cqw_0.5cqw_rgba(0,0,0,0.25)] ${className}`} style={{ background: tone(type), width: "2.4cqw", height: "2.4cqw", fontSize: "1.2cqw", fontFamily: "var(--s-title)" }}>
      {n}
    </span>
  );
}

/**
 * Tutorial page: a panel of numbered explanations on the left; on the right, the live Power BI page
 * where each explained visual is framed in its item's color with its number on the corner (frames
 * placed from the report's kept schema).
 */
export const tutorial = defineTemplate({
  id: "tutorial",
  description:
    "Tutorial of a report page: numbered explanations on the left, the Power BI page on the right where each explained visual (or group of visuals) is framed with its item's number — one number per visual: never two items on the same visual, put what is said about a visual in its one item. report + page from the context; items: text (markdown), type (default | info), visual (a visual id of that page) or visuals (several ids: one frame around all of them); an item without a visual (a general tip) gets a dot instead of a number.",
  args: z.object({
    title: z.string(),
    panel: z.string().optional().describe("Small heading of the panel, e.g. \"Sales page — top of the page\""),
    report: z.string().describe("Report id"),
    page: z.string().optional().describe("Page name (id)"),
    crop: cropSchema.optional().describe("Only a zone of the page (1280×720 px), e.g. its top half"),
    items: z.array(item).min(1).max(8),
  }),
  Component: ({ args }) => {
    const entry = useReportRef(args.report);
    const size = fitText(args.items.map((i) => i.text), 22 * 36, { max: 1.35, min: 0.85 });
    const numbers = numbering(args.items);
    return (
      <Frame title={args.title}>
        <div className="grid h-full grid-cols-[30fr_70fr] gap-[2.4cqw]">
          <M.div variants={stagger(0.1, 0.2)} className="flex min-h-0 flex-col gap-[1.4cqw] overflow-hidden rounded-[0.4cqw] border px-[1.6cqw] py-[1.4cqw] shadow-[0_0.3cqw_1cqw_rgba(0,0,0,0.08)]" style={{ borderColor: "color-mix(in srgb, var(--s-text) 18%, transparent)", background: "color-mix(in srgb, var(--s-text) 2%, var(--s-bg))", fontSize: `${size}cqw` }}>
            {args.panel && (
              <M.p variants={rise} className="text-[1.05cqw] font-medium tracking-wide uppercase" style={{ color: "var(--s-secondary)" }}>
                {args.panel}
              </M.p>
            )}
            {args.items.map((it, i) => (
              <M.div key={i} variants={rise} className="flex gap-[1cqw] leading-snug">
                <Badge n={numbers[i].n} type={it.type} />
                <div className="min-w-0 flex-1 pt-[0.2cqw]">
                  <SlideMarkdown text={it.text} />
                </div>
              </M.div>
            ))}
          </M.div>
          <M.div variants={{ hidden: { opacity: 0, x: "2cqw" }, show: { opacity: 1, x: 0, transition: { duration: 0.9, ease } } }} className="min-h-0">
            <PowerBiFrame
              report={args.report}
              page={args.page}
              crop={args.crop}
              overlay={(place) =>
                args.items.map((it, i) => {
                  const found = unionZone(numbers[i].visuals.map((v) => visualZone(entry, args.page, v)));
                  if (!found || numbers[i].n === undefined) return null;
                  const zone = { x: found.x - PAD.x, y: found.y - PAD.y, w: found.w + 2 * PAD.x, h: found.h + 2 * PAD.y };
                  return (
                    <M.div key={i} style={place(zone)} initial={{ opacity: 0, scale: 1.04 }} animate={{ opacity: 1, scale: 1 }} transition={{ duration: 0.6, delay: 0.9 + i * 0.25, ease }} className="pointer-events-none">
                      <div className="absolute inset-0 rounded-[0.3cqw] border-[0.22cqw]" style={{ borderColor: tone(it.type) }} />
                      <Badge n={numbers[i].n} type={it.type} className="absolute -top-[1.2cqw] -right-[1.2cqw]" />
                    </M.div>
                  );
                })
              }
            />
          </M.div>
        </div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    if (args.panel) slide.addText(args.panel.toUpperCase(), { x: 0.6, y: 2.0, w: 3.6, h: 0.35, fontSize: 10, color: hex(ctx.theme.secondary ?? "#8b5a2b"), fontFace: ctx.theme.fontText });
    const h = Math.min(1.2, 4.4 / args.items.length);
    const numbers = numbering(args.items);
    args.items.forEach((it, i) => {
      const y = 2.45 + i * h;
      const n = numbers[i].n;
      const d = n === undefined ? 0.14 : 0.34;
      slide.addShape("ellipse", { x: 0.6 + (0.34 - d) / 2, y: y + (0.34 - d) / 2, w: d, h: d, fill: { color: hex(it.type === "info" ? (ctx.theme.success ?? "#15803d") : (ctx.theme.secondary ?? "#8b5a2b")) }, line: { type: "none" } as never });
      if (n !== undefined) slide.addText(String(n), { x: 0.6, y, w: 0.34, h: 0.34, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
      slide.addText(plain(it.text), { x: 1.05, y: y - 0.05, w: 3.2, h: h - 0.05, fontSize: 11, valign: "top", fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    });
    pptxPowerBi(slide, ctx.theme, args, { x: 4.5, y: 2.0, w: 8.2, h: 4.9 });
  },
});
