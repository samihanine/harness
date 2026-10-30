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
      <Frame title={args.title}>
        <div className="flex h-full flex-col gap-[1.4cqw]">
          {args.lead && (
            <M.p variants={rise} className="max-w-[70cqw] text-[1.4cqw] leading-snug opacity-80">
              <Rich text={args.lead} />
            </M.p>
          )}
          <M.div variants={rise} className="flex min-h-0 flex-1 items-center justify-center">
            <table className="max-h-full border-collapse overflow-hidden rounded-[0.6cqw]" style={{ fontSize: `${size}cqw`, width: `${widths.reduce((a, b) => a + b, 0)}cqw` }}>
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
    pptxTitle(slide, args.title, ctx);
    const { widths, size } = layout(args.columns, args.rows);
    const scale = 11.7 / Math.max(widths.reduce((a, b) => a + b, 0), 1);
    const fill = (r: number, c: number) => args.cellColors?.find((x) => x.row === r + 1 && x.column === c + 1)?.color ?? args.columnColors?.find((x) => x.column === c + 1)?.color ?? args.rowColors?.find((x) => x.row === r + 1)?.color;
    const header = args.columns.map((c) => ({ text: plain(c), options: { bold: true, color: "FFFFFF", fill: { color: hex(args.headerColor ?? ctx.theme.primary) } } }));
    const body = args.rows.map((row, r) => args.columns.map((_, c) => ({ text: plain(String(row[c] ?? "")), options: { align: (typeof row[c] === "number" ? "right" : "left") as "right" | "left", ...(fill(r, c) && { fill: { color: hex(fill(r, c)!) } }) } })));
    slide.addTable([header, ...body], { x: 0.8, y: args.lead ? 2.8 : 2.2, colW: widths.map((w) => w * scale), fontSize: Math.round(size * 9.5), fontFace: ctx.theme.fontText, color: hex(ctx.theme.text), border: { type: "solid", color: "E5E7EB", pt: 0.5 } });
    if (args.lead) slide.addText(plain(args.lead), { x: 0.8, y: 2.1, w: 11.5, h: 0.6, fontSize: 15, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
  },
});

/**
 * Column widths (cqw) from each column's longest text, and the text size (cqw) that fills most of the
 * slide without overflowing it: smaller with more rows, more columns or longer texts. A narrow table is
 * widened to most of the slide's width (the extra space shared by the columns).
 */
function layout(columns: string[], rows: (string | number)[][]) {
  // A long text wraps: a column is at most WRAP characters wide, its rows taking more lines.
  const WRAP = 30;
  const len = (v: unknown) => plain(String(v ?? "")).length;
  const longest = columns.map((c, i) => Math.min(WRAP, Math.max(len(c), ...rows.map((r) => len(r[i])), 3)));
  const lines = [columns, ...rows].reduce((n, r) => n + Math.max(1, ...r.map((v, i) => Math.ceil(len(v) / (longest[i] ?? WRAP)))), 0);
  const chars = longest.reduce((a, b) => a + b, 0);
  // Height: the lines of the header and rows in ~38cqw of usable height; width: ~88cqw for all the characters.
  const byHeight = 38 / (lines * 1.45 + (rows.length + 1) * 0.65);
  const byWidth = 88 / (chars * 0.62 + columns.length * 1.8);
  const size = Math.max(0.8, Math.min(2, byHeight, byWidth));
  const natural = longest.map((n) => n * size * 0.62 + size * 1.8);
  const total = natural.reduce((a, b) => a + b, 0);
  const target = Math.min(88, Math.max(total, 70));
  return { widths: natural.map((w) => (w * target) / total), size };
}
