/** Reading the workbook: its sheets, the values Excel computes (formulas included), and a picture of a sheet. */
import { z } from "zod";
import { excel, rangePath, sheetPath } from "@server/workbook/excel";
import { sheetHtml, sheetPreview } from "@server/workbook/preview";
import { renderHtml } from "@server/capture/browser";
import { base64, tool } from "../tool";
import { gridText, workbookOf } from "./shared";

export const workbookInfo = tool({
  name: "workbook_info",
  description: "The workbook: its sheets (used range, size), tables, charts and the models' connections (for CUBE formulas). Start with it.",
  args: z.object({}),
  async run(_, ctx) {
    const id = workbookOf(ctx);
    const sheets = (await excel(id, "/worksheets?$select=name,position")).value as { name: string; position: number }[];
    const lines = await Promise.all(
      sheets.sort((a, b) => a.position - b.position).map(async (s) => {
        const [used, tables, charts] = await Promise.all([
          excel(id, `${sheetPath(s.name)}/usedRange(valuesOnly=true)?$select=address,rowCount,columnCount`).catch(() => undefined),
          excel(id, `${sheetPath(s.name)}/tables?$select=name`).catch(() => ({ value: [] })),
          excel(id, `${sheetPath(s.name)}/charts?$select=name`).catch(() => ({ value: [] })),
        ]);
        return `- "${s.name}": ${used?.address ? `${used.address.split("!")[1]} (${used.rowCount} rows × ${used.columnCount} columns)` : "empty"}${tables.value.length ? `; tables ${tables.value.map((t: { name: string }) => t.name).join(", ")}` : ""}${charts.value.length ? `; charts ${charts.value.map((c: { name: string }) => c.name).join(", ")}` : ""}`;
      }),
    );
    const models = (ctx.scope.datasets ?? []).map((d) => `"PBI ${d.name}" (model "${d.name}")`);
    return [`Sheets:`, ...lines, models.length ? `Connections for CUBE formulas: ${models.join(", ")}` : "No model connected."].join("\n");
  },
});

export const readRange = tool({
  name: "read_range",
  description: "The values of a range as Excel computes them (formulas, CUBE formulas on the models included); formulas: true adds each cell's formula. At most 200 rows × 40 columns: read big sheets in parts, or summarise them with formulas.",
  args: z.object({ sheet: z.string(), range: z.string().describe("e.g. A1:F30"), formulas: z.boolean().optional() }),
  async run(a, ctx) {
    const id = workbookOf(ctx);
    const r = await excel(id, `${rangePath(a.sheet, a.range)}?$select=address,values,formulas,rowCount,columnCount,rowIndex,columnIndex`);
    if (r.rowCount > 200 || r.columnCount > 40) throw new Error(`${r.address} is ${r.rowCount} × ${r.columnCount}: read at most 200 rows × 40 columns at once.`);
    return `${r.address}:\n${gridText(r.values, a.formulas ? r.formulas : undefined, r.rowIndex + 1, r.columnIndex + 1)}`;
  },
});

export const lookAtSheet = tool({
  name: "look_at_sheet",
  description: "A picture of a sheet as the user sees it in the app (first rows and columns, with formats and conditional formats): to check the look after formatting.",
  args: z.object({ sheet: z.string().optional() }),
  async run(a, ctx) {
    ctx.status("Drawing the sheet…");
    const p = await sheetPreview(workbookOf(ctx), a.sheet);
    const png = await renderHtml(sheetHtml(p), 1400);
    return { text: `Sheet "${p.sheet}" (${p.total.rows} rows × ${p.total.cols} columns${p.total.rows > p.cells.length || p.total.cols > p.widths.length ? `; shown: ${p.cells.length} × ${p.widths.length}` : ""}).`, images: [{ mediaType: "image/png", data: base64(png) }] };
  },
});
