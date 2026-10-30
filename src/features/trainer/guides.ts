/**
 * The guides of the report viewer (Trainer): one markdown guide per visual of a report, kept in this
 * browser (IndexedDB), downloaded as an Excel file and imported from Excel or CSV (Settings › Guides).
 */
import { readSheet } from "read-excel-file/browser";
import writeXlsxFile from "write-excel-file/browser";
import type { ReportSchema } from "@/services/report-schema";
import { all, put, type GuideEntry } from "@/lib/store";

/** The columns of the Excel / CSV file, in order. */
export const GUIDE_COLUMNS = ["id", "report_id", "page", "visual", "title", "content"] as const;

const nextId = (guides: GuideEntry[]) => `GUIDE_${String(guides.reduce((n, g) => Math.max(n, Number(g.id.match(/(\d+)$/)?.[1] ?? 0)), 0) + 1).padStart(3, "0")}`;

/** Writes a guide: an existing one (same id, or same report, page and visual) is updated. */
export async function saveGuide(given: Omit<GuideEntry, "id" | "updatedAt"> & { id?: string }) {
  const guides = await all("guides");
  const old = guides.find((g) => (given.id && g.id === given.id) || (g.reportId === given.reportId && g.page === given.page && g.visual === given.visual));
  return put("guides", { ...old, ...given, id: old?.id ?? given.id ?? nextId(guides), updatedAt: Date.now() });
}

/** Visuals without data of their own (decoration, navigation): no guide. */
const DECORATIONS = ["shape", "basicShape", "image", "textbox", "actionButton", "pageNavigator", "bookmarkNavigator"];

/**
 * An empty guide for each data visual of a report that has none yet (when the report is added), so the
 * downloaded Excel file lists every visual, ready to be filled.
 */
export async function addEmptyGuides(reportId: string, schema: ReportSchema) {
  const guides = await all("guides");
  const has = (page: string, visual: string) => guides.some((g) => g.reportId === reportId && g.page === page && g.visual === visual);
  let added = 0;
  for (const v of schema.visuals.filter((x) => !x.hidden && !DECORATIONS.includes(x.type))) {
    if (has(v.page, v.name)) continue;
    const guide: GuideEntry = { id: nextId(guides), reportId, page: v.page, visual: v.name, title: v.title ?? "", content: "", updatedAt: Date.now() };
    guides.push(guide);
    await put("guides", guide);
    added++;
  }
  return added;
}

/** Every guide as an Excel file (one row per guide). */
export async function downloadGuides() {
  const guides = await all("guides");
  const header = GUIDE_COLUMNS.map((c) => ({ value: c, fontWeight: "bold" as const }));
  const rows = guides.map((g) => [g.id, g.reportId, g.page, g.visual, g.title, g.content].map((value) => ({ type: String, value })));
  await writeXlsxFile([header, ...rows]).toFile(`guides-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/** The cells of a CSV text (quotes handled; the delimiter is the one splitting the header most). */
function csvRows(text: string) {
  const first = text.slice(0, text.indexOf("\n") >>> 0 || undefined);
  const d = [",", ";", "\t"].sort((a, b) => first.split(b).length - first.split(a).length)[0];
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') (cell += '"'), i++;
      else if (c === '"') quoted = false;
      else cell += c;
    } else if (c === '"') quoted = true;
    else if (c === d) row.push(cell), (cell = "");
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(cell);
      if (row.some((x) => x !== "")) rows.push(row);
      (row = []), (cell = "");
    } else cell += c;
  }
  row.push(cell);
  if (row.some((x) => x !== "")) rows.push(row);
  return rows;
}

/** Imports guides from an Excel (.xlsx) or CSV file with the columns of GUIDE_COLUMNS; returns how many were written. */
export async function importGuides(file: File) {
  const rows: unknown[][] = /\.csv$/i.test(file.name) ? csvRows(await file.text()) : ((await readSheet(file)) as unknown as unknown[][]);
  const header = (rows[0] ?? []).map((h) => String(h ?? "").trim().toLowerCase());
  const at = (name: string) => header.indexOf(name);
  for (const c of ["report_id", "page", "visual", "title"]) if (at(c) < 0) throw new Error(`Column "${c}" is missing (columns: ${GUIDE_COLUMNS.join(", ")}).`);
  let count = 0;
  for (const r of rows.slice(1)) {
    const cell = (name: string) => (at(name) >= 0 ? String(r[at(name)] ?? "").trim() : "");
    if (!cell("report_id") || !cell("page") || !cell("visual")) continue;
    await saveGuide({ id: cell("id") || undefined, reportId: cell("report_id"), page: cell("page"), visual: cell("visual"), title: cell("title"), content: at("content") >= 0 ? String(r[at("content")] ?? "") : "" });
    count++;
  }
  return count;
}
