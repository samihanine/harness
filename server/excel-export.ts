/** An analysis as an Excel workbook: the findings, each block's rows, each source's rows, the SQL of every step. */
import writeXlsxFile from "write-excel-file/node";
import { pivotRows, type Analysis, type Rows } from "@shared/analysis";

const sheetName = (s: string, used: Set<string>) => {
  let base = s.replace(/[\\/?*[\]:]/g, " ").trim().slice(0, 28) || "Sheet";
  let name = base;
  for (let n = 2; used.has(name.toLowerCase()); n++) name = `${base.slice(0, 25)} ${n}`;
  used.add(name.toLowerCase());
  return name;
};

const header = (t: string) => ({ value: t, fontWeight: "bold" as const, backgroundColor: "#F3F4F6" });
const cell = (v: unknown) => ({ value: v === null || v === undefined ? null : typeof v === "number" || typeof v === "boolean" ? v : String(v), ...(typeof v === "number" ? { type: Number } : {}) });

function table(rows: Rows, columns?: string[]) {
  const cols = columns ?? [...new Set(rows.flatMap((r) => Object.keys(r)))];
  return [cols.map(header), ...rows.map((r) => cols.map((c) => cell(r[c])))];
}

export async function analysisToExcel(a: Analysis): Promise<Buffer> {
  const used = new Set<string>();
  const sheets: { data: any[][]; sheet: string; columns?: { width: number }[] }[] = [];
  const summary: any[][] = [[{ value: a.title, fontWeight: "bold", fontSize: 14 }], []];
  for (const b of a.blocks) {
    if (b.type === "text") for (const line of b.markdown.split("\n")) summary.push([{ value: line.replace(/[*_`#>]/g, "") }]);
    else if (b.type === "kpis") for (const it of b.items) summary.push([{ value: it.label }, cell(b.rows?.[0]?.[it.column])]);
    else summary.push([{ value: `${b.title ?? b.type}: see sheet`, textColor: "#6B7280" }]);
  }
  sheets.push({ sheet: sheetName("Analysis", used), data: summary, columns: [{ width: 60 }, { width: 18 }] });
  for (const b of a.blocks) {
    if (b.type === "text" || b.type === "kpis" || !b.rows) continue;
    const data = b.type === "pivot" ? (() => { const p = pivotRows(b.rows!, b.pivot); return table(p.lines as Rows, [...b.pivot.rows, ...p.columns]); })() : table(b.rows, b.type === "table" ? b.columns : undefined);
    sheets.push({ sheet: sheetName(b.title ?? b.type, used), data });
  }
  const steps: any[][] = [[header("Name"), header("SQL"), header("Description")], ...a.steps.map((s) => [{ value: s.name }, { value: s.sql }, { value: s.description ?? "" }]), [], [header("Source"), header("Kind"), header("Query / file")], ...a.sources.map((s) => [{ value: s.name }, { value: s.kind }, { value: s.kind === "dax" ? `${s.datasetName}: ${s.query}` : `${s.fileName}${s.sheet ? ` / ${s.sheet}` : ""}` }])];
  sheets.push({ sheet: sheetName("Method", used), data: steps, columns: [{ width: 24 }, { width: 80 }, { width: 40 }] });
  return (await writeXlsxFile(sheets.map((s) => ({ data: s.data, sheet: s.sheet, columns: s.columns })) as never, {} as never)) as unknown as Buffer;
}
