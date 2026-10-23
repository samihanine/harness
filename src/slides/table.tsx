import { z } from "zod";
import { Frame, M, Rich, defineTemplate, hex, plain, pptxTitle, rise } from "./kit";

const color = z.string().regex(/^#[0-9a-fA-F]{6}$/, "a color #RRGGBB");

/**
 * A table, centered: column widths follow their longest text, the text size follows the number of
 * rows and columns (never overflowing the slide); colors are optional, for the header, whole rows,
 * whole columns or single cells (a cell's color wins over its column's, which wins over its row's).
 */
export const table = defineTemplate({
  id: "table",
  description:
    "A table centered on the page (sized automatically to fit): columns (header texts), rows (cell texts, **bold** / ==highlight== allowed), optional colors: headerColor, rowColors (by row number from 1), columnColors (by column number from 1), cellColors [{row, column, color}], lead sentence under the title.",
  args: z.object({
    title: z.string(),
    eyebrow: z.string().optional(),
    lead: z.string().optional(),
    columns: z.array(z.string()).min(1).max(10),
    rows: z.array(z.array(z.union([z.string(), z.number()]))).min(1).max(18),
    headerColor: color.optional().describe("Header background (default: the accent color)"),
    rowColors: z.array(z.object({ row: z.number(), color })).optional(),
    columnColors: z.array(z.object({ column: z.number(), color })).optional(),
    cellColors: z.array(z.object({ row: z.number(), column: z.number(), color })).optional(),
  }),
  Component: ({ args }) => {
    const { widths, size } = layout(args.columns, args.rows);
    const fill = (r: number, c: number) => args.cellColors?.find((x) => x.row === r + 1 && x.column === c + 1)?.color ?? args.columnColors?.find((x) => x.column === c + 1)?.color ?? args.rowColors?.find((x) => x.row === r + 1)?.color;
    return (
      <Frame eyebrow={args.eyebrow} title={args.title}>
        <div className="flex h-full flex-col gap-[1.4cqw]">
          {args.lead && (
            <M.p variants={rise} className="max-w-[70cqw] text-[1.4cqw] leading-snug opacity-80">
              <Rich text={args.lead} />
            </M.p>
          )}
          <M.div variants={rise} className="flex min-h-0 flex-1 items-center justify-center">
            <table className="max-h-full border-collapse overflow-hidden rounded-[0.6cqw]" style={{ fontSize: `${size}cqw`, width: `${Math.min(88, widths.reduce((a, b) => a + b, 0))}cqw` }}>
              <colgroup>
                {widths.map((w, i) => (
                  <col key={i} style={{ width: `${w}cqw` }} />
                ))}
              </colgroup>
              <thead>
                <tr>
                  {args.columns.map((c, i) => (
                    <th key={i} className="px-[0.9em] py-[0.55em] text-left font-semibold" style={{ background: args.headerColor ?? "var(--s-primary)", color: "#fff" }}>
                      <Rich text={c} />
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {args.rows.map((row, r) => (
                  <tr key={r} style={{ background: r % 2 ? "color-mix(in srgb, var(--s-text) 4%, var(--s-bg))" : "var(--s-bg)" }}>
                    {args.columns.map((_, c) => (
                      <td key={c} className="border-b px-[0.9em] py-[0.5em] align-middle" style={{ borderColor: "color-mix(in srgb, var(--s-text) 10%, transparent)", background: fill(r, c), textAlign: typeof row[c] === "number" ? "right" : "left", fontVariantNumeric: "tabular-nums" }}>
                        <Rich text={String(row[c] ?? "")} />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </M.div>
        </div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx, args.eyebrow);
    const { widths, size } = layout(args.columns, args.rows);
    const scale = 11.7 / Math.max(widths.reduce((a, b) => a + b, 0), 1);
    const fill = (r: number, c: number) => args.cellColors?.find((x) => x.row === r + 1 && x.column === c + 1)?.color ?? args.columnColors?.find((x) => x.column === c + 1)?.color ?? args.rowColors?.find((x) => x.row === r + 1)?.color;
    const header = args.columns.map((c) => ({ text: plain(c), options: { bold: true, color: "FFFFFF", fill: { color: hex(args.headerColor ?? ctx.theme.primary) } } }));
    const body = args.rows.map((row, r) => args.columns.map((_, c) => ({ text: plain(String(row[c] ?? "")), options: { align: (typeof row[c] === "number" ? "right" : "left") as "right" | "left", ...(fill(r, c) && { fill: { color: hex(fill(r, c)!) } }) } })));
    slide.addTable([header, ...body], { x: 0.8, y: args.lead ? 2.8 : 2.2, colW: widths.map((w) => w * scale), fontSize: Math.round(size * 9), fontFace: ctx.theme.fontText, color: hex(ctx.theme.text), border: { type: "solid", color: "E5E7EB", pt: 0.5 } });
    if (args.lead) slide.addText(plain(args.lead), { x: 0.8, y: 2.1, w: 11.5, h: 0.6, fontSize: 15, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
  },
});

/**
 * Column widths (cqw) from each column's longest text, and the text size (cqw) that keeps the table
 * inside the slide: smaller with more rows, more columns or longer texts.
 */
function layout(columns: string[], rows: (string | number)[][]) {
  const longest = columns.map((c, i) => Math.max(plain(c).length, ...rows.map((r) => plain(String(r[i] ?? "")).length), 3));
  const chars = longest.reduce((a, b) => a + b, 0);
  // Height: header + rows in ~30cqw of usable height; width: ~84cqw for all the characters.
  const byHeight = 30 / ((rows.length + 1) * 2.1);
  const byWidth = 84 / (chars * 0.62 + columns.length * 1.8);
  const size = Math.max(0.75, Math.min(1.5, byHeight, byWidth));
  return { widths: longest.map((n) => n * size * 0.62 + size * 1.8), size };
}
