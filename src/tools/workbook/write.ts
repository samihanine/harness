/**
 * Changing the workbook through Excel: values and formulas, formats, conditional formats, validation lists,
 * sheets, tables, charts. Excel computes everything (also over hundreds of thousands of rows).
 */
import { z } from "zod";
import { excel, rangePath, sheetPath } from "@server/workbook/excel";
import { addConditionalFormat, setValidation, type Rule } from "@server/workbook/ooxml";
import { tool } from "../tool";
import { blockAt, letters, workbookOf } from "./shared";

const cellValue = z.union([z.string(), z.number(), z.boolean(), z.null()]);

export const writeRange = tool({
  name: "write_range",
  description:
    "Write a block of cells from `at` (its top-left cell): rows of values; a text starting with = is a formula (Excel's English names and , separators: =SUMIFS(Data!C:C,Data!A:A,A2), =XLOOKUP(…), =CUBEVALUE(\"PBI Sales\",…)). Excel computes them: read_range shows the results. Big data: prefer formulas over a sheet already in the workbook to copying values.",
  args: z.object({ sheet: z.string(), at: z.string().describe("Top-left cell, e.g. A1"), values: z.array(z.array(cellValue)).min(1) }),
  changes: true,
  async run(a, ctx) {
    const id = workbookOf(ctx);
    const cols = Math.max(...a.values.map((r) => r.length));
    const rows = a.values.map((r) => [...r, ...Array(cols - r.length).fill(null)].map((v) => (v === null ? "" : v)));
    // In parts: Excel's API takes a few MB per call.
    const CHUNK = Math.max(1, Math.floor(20000 / cols));
    const start = Number(a.at.replace(/\$/g, "").match(/\d+$/)?.[0] ?? 1);
    const col = a.at.replace(/\$/g, "").match(/^[A-Z]+/i)?.[0] ?? "A";
    for (let i = 0; i < rows.length; i += CHUNK) {
      const part = rows.slice(i, i + CHUNK);
      await excel(id, rangePath(a.sheet, blockAt(`${col}${start + i}`, part.length, cols)), "PATCH", { formulas: part });
    }
    const errors = (await excel(id, `${rangePath(a.sheet, blockAt(a.at, Math.min(rows.length, 200), cols))}?$select=values`)).values.flatMap((r: unknown[], i: number) => r.map((v, j) => (typeof v === "string" && /^#(N\/A|VALUE!|REF!|NAME\?|DIV\/0!|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)/.test(v) ? `row ${i + 1} col ${j + 1}: ${v}` : undefined)).filter(Boolean));
    return `Wrote ${rows.length} × ${cols} at ${a.sheet}!${blockAt(a.at, rows.length, cols)}.${errors.length ? ` Cells with errors (offset from ${a.at}): ${errors.slice(0, 20).join(", ")} — fix them.` : ""}`;
  },
});

const sides = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];

export const formatRange = tool({
  name: "format_range",
  description:
    'Format a range: {"bold": true, "italic": true, "fontColor": "#FFFFFF", "fontSize": 12, "fontName": "Calibri", "fill": "#1C4358", "numberFormat": "#,##0;[Red]-#,##0" | "0.0%" | "dd/mm/yyyy", "horizontalAlignment": "Center", "verticalAlignment": "Center", "wrapText": true, "indent": 1, "columnWidth": 18 (characters), "rowHeight": 24 (points), "autofit": true, "borders": {"style": "Continuous" | "Dash" | "Double", "weight": "Thin" | "Medium" | "Thick", "color": "#D1D5DB", "where": "all" | "outside" | "bottom" | "top" | "inside"}, "merge": true}.',
  args: z.object({
    sheet: z.string(),
    range: z.string(),
    format: z.object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      fontColor: z.string().optional(),
      fontSize: z.number().optional(),
      fontName: z.string().optional(),
      fill: z.string().optional(),
      numberFormat: z.string().optional(),
      horizontalAlignment: z.enum(["General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection"]).optional(),
      verticalAlignment: z.enum(["Top", "Center", "Bottom", "Justify"]).optional(),
      wrapText: z.boolean().optional(),
      indent: z.number().optional(),
      columnWidth: z.number().optional(),
      rowHeight: z.number().optional(),
      autofit: z.boolean().optional(),
      borders: z.object({ style: z.string().optional(), weight: z.string().optional(), color: z.string().optional(), where: z.enum(["all", "outside", "inside", "top", "bottom", "left", "right", "none"]).optional() }).optional(),
      merge: z.boolean().optional(),
    }),
  }),
  changes: true,
  async run(a, ctx) {
    const id = workbookOf(ctx);
    const r = rangePath(a.sheet, a.range);
    const f = a.format;
    const font = { ...(f.bold !== undefined && { bold: f.bold }), ...(f.italic !== undefined && { italic: f.italic }), ...(f.underline !== undefined && { underline: f.underline ? "Single" : "None" }), ...(f.fontColor && { color: f.fontColor }), ...(f.fontSize && { size: f.fontSize }), ...(f.fontName && { name: f.fontName }) };
    if (Object.keys(font).length) await excel(id, `${r}/format/font`, "PATCH", font);
    if (f.fill) await excel(id, `${r}/format/fill`, "PATCH", { color: f.fill });
    const format = { ...(f.horizontalAlignment && { horizontalAlignment: f.horizontalAlignment }), ...(f.verticalAlignment && { verticalAlignment: f.verticalAlignment }), ...(f.wrapText !== undefined && { wrapText: f.wrapText }), ...(f.columnWidth && { columnWidth: f.columnWidth * 7 + 5 }), ...(f.rowHeight && { rowHeight: f.rowHeight }), ...(f.indent !== undefined && { indentLevel: f.indent }) };
    if (Object.keys(format).length) await excel(id, `${r}/format`, "PATCH", format);
    if (f.numberFormat) {
      const size = await excel(id, `${r}?$select=rowCount,columnCount`);
      await excel(id, r, "PATCH", { numberFormat: Array.from({ length: size.rowCount }, () => Array(size.columnCount).fill(f.numberFormat)) });
    }
    if (f.autofit) (await excel(id, `${r}/format/autofitColumns`, "POST", {}), await excel(id, `${r}/format/autofitRows`, "POST", {}));
    if (f.borders) {
      const where = f.borders.where ?? "all";
      const list = where === "all" ? sides : where === "outside" ? sides.slice(0, 4) : where === "inside" ? sides.slice(4) : where === "none" ? sides : [`Edge${where[0].toUpperCase()}${where.slice(1)}`];
      for (const side of list) await excel(id, `${r}/format/borders('${side}')`, "PATCH", where === "none" ? { style: "None" } : { style: f.borders.style ?? "Continuous", weight: f.borders.weight ?? "Thin", color: f.borders.color ?? "#000000" }).catch(() => undefined);
    }
    if (f.merge !== undefined) await excel(id, `${r}/${f.merge ? "merge" : "unmerge"}`, "POST", f.merge ? { across: false } : {});
    return `Formatted ${a.sheet}!${a.range}.`;
  },
});

export const conditionalFormat = tool({
  name: "conditional_format",
  description:
    'Conditional formatting of a range (colors by content): rule "cellValue" (operator GreaterThan / LessThan / Between / EqualTo / NotEqualTo / GreaterThanOrEqual / LessThanOrEqual, formula1 [, formula2] e.g. "0", fill / fontColor / bold), "text" (contains `text`), "custom" (a formula true for the cell, relative to the range\'s first cell, e.g. "=$C2<$D2"), "colorScale" (colors: 2 or 3, low to high, e.g. ["#F8696B","#FFEB84","#63BE7B"]), "dataBar" (fill). clear: true removes the range\'s rules first.',
  args: z.object({
    sheet: z.string(),
    range: z.string(),
    rule: z.enum(["cellValue", "text", "custom", "colorScale", "dataBar"]),
    operator: z.string().optional(),
    formula1: z.string().optional(),
    formula2: z.string().optional(),
    text: z.string().optional(),
    formula: z.string().optional(),
    colors: z.array(z.string()).optional(),
    fill: z.string().optional(),
    fontColor: z.string().optional(),
    bold: z.boolean().optional(),
    clear: z.boolean().optional(),
  }),
  changes: true,
  async run(a, ctx) {
    // Written in the file (Excel's API has no conditional formats): Excel applies them, the app shows them.
    const rule: Rule = a.rule === "cellValue" ? { rule: "cellValue", operator: a.operator ?? "GreaterThan", formula1: a.formula1 ?? "0", formula2: a.formula2 } : a.rule === "text" ? { rule: "text", text: a.text ?? "" } : a.rule === "custom" ? { rule: "custom", formula: a.formula ?? "TRUE" } : a.rule === "colorScale" ? { rule: "colorScale", colors: a.colors ?? [] } : { rule: "dataBar", fill: a.fill };
    await addConditionalFormat(workbookOf(ctx), a.sheet, a.range, rule, { fill: a.fill, fontColor: a.fontColor, bold: a.bold }, a.clear);
    return `Conditional format (${a.rule}) on ${a.sheet}!${a.range}.`;
  },
});

export const dataValidation = tool({
  name: "data_validation",
  description: 'A drop-down list on cells: list (values) or source (a range, e.g. "=Lists!$A$2:$A$20"); clear: true removes the validation.',
  args: z.object({ sheet: z.string(), range: z.string(), list: z.array(z.string()).optional(), source: z.string().optional(), message: z.string().optional(), clear: z.boolean().optional() }),
  changes: true,
  async run(a, ctx) {
    if (!a.clear && !a.list?.length && !a.source) throw new Error("Give list or source.");
    await setValidation(workbookOf(ctx), a.sheet, a.range, a.clear ? undefined : { values: a.list, source: a.source, message: a.message });
    return a.clear ? "Validation removed." : `Drop-down list on ${a.sheet}!${a.range}.`;
  },
});

export const sheetAction = tool({
  name: "sheet",
  description: "Sheets: add (name), rename (to), delete, move (position from 0), freeze (rows and/or columns at the top-left; 0 to unfreeze), tab color.",
  args: z.object({ action: z.enum(["add", "rename", "delete", "move", "freeze", "color"]), sheet: z.string(), to: z.string().optional(), position: z.number().int().optional(), rows: z.number().int().optional(), columns: z.number().int().optional(), color: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const id = workbookOf(ctx);
    const s = sheetPath(a.sheet);
    if (a.action === "add") await excel(id, "/worksheets/add", "POST", { name: a.sheet });
    else if (a.action === "rename") await excel(id, s, "PATCH", { name: a.to });
    else if (a.action === "delete") await excel(id, s, "DELETE");
    else if (a.action === "move") await excel(id, s, "PATCH", { position: a.position ?? 0 });
    else if (a.action === "color") await excel(id, s, "PATCH", { tabColor: a.color ?? "#B39251" });
    else {
      await excel(id, `${s}/freezePanes/unfreeze`, "POST", {}).catch(() => undefined);
      if (a.rows && a.columns) await excel(id, `${s}/freezePanes/freezeAt`, "POST", { frozenRange: `A1:${letters(a.columns)}${a.rows}` });
      else if (a.rows) await excel(id, `${s}/freezePanes/freezeRows`, "POST", { count: a.rows });
      else if (a.columns) await excel(id, `${s}/freezePanes/freezeColumns`, "POST", { count: a.columns });
    }
    return `Sheet "${a.sheet}": ${a.action} done.`;
  },
});

export const createTable = tool({
  name: "create_table",
  description: "Turn a range with a header row into an Excel table (filters, banded rows, structured references like Sales[Amount] in formulas): name and style (e.g. TableStyleMedium2, TableStyleLight9).",
  args: z.object({ sheet: z.string(), range: z.string(), name: z.string().optional(), style: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const id = workbookOf(ctx);
    const t = await excel(id, `${sheetPath(a.sheet)}/tables/add`, "POST", { address: a.range, hasHeaders: true });
    await excel(id, `/tables/${encodeURIComponent(t.id)}`, "PATCH", { ...(a.name && { name: a.name.replace(/[^\w]/g, "_") }), ...(a.style && { style: a.style }) });
    return `Table ${a.name ?? t.name} on ${a.sheet}!${a.range}.`;
  },
});

export const addChart = tool({
  name: "add_chart",
  description: "A chart from a range (first column the categories, then one series per column): type ColumnClustered, BarClustered, Line, LineMarkers, Pie, Doughnut, Area, ColumnStacked…; placed over `at` (e.g. H2:P20).",
  args: z.object({ sheet: z.string(), type: z.string(), source: z.string(), at: z.string(), title: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const id = workbookOf(ctx);
    const c = await excel(id, `${sheetPath(a.sheet)}/charts/add`, "POST", { type: a.type, sourceData: a.source, seriesBy: "Auto" });
    const chart = `${sheetPath(a.sheet)}/charts/${encodeURIComponent(c.id ?? c.name)}`;
    await excel(id, `${chart}/setPosition`, "POST", { startCell: a.at.split(":")[0], endCell: a.at.split(":")[1] ?? a.at }).catch(() => undefined);
    if (a.title) await excel(id, `${chart}/title`, "PATCH", { text: a.title, visible: true }).catch(() => undefined);
    return `Chart ${a.type} on ${a.sheet} (${a.at}).`;
  },
});
