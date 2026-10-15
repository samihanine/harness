/**
 * Publishing a research: its workbook is written by the app (workbook.ts) into the Excel file chosen
 * by link (its content replaced each time), then Excel for the web refreshes its PivotTables from the
 * semantic models and draws its charts (Graph workbook API).
 */
import { uploadFile, writeItem } from "@/server/ms";
import { resolveExcel, toBase64 } from "./excel";
import { graph } from "./ms";
import { patch, type ResearchEntry } from "./store";
import { buildWorkbook, sheetName } from "./workbook";

let tenant: Promise<string> | undefined;
/** The organization's id (the pivots' connections name it). */
const tenantId = () => (tenant ??= graph("/organization?$select=id").then((r) => r.value[0].id as string));

const FOLDER = "Hub/Researcher";
const book = (f: { driveId: string; itemId: string }) => `/drives/${f.driveId}/items/${f.itemId}/workbook`;
const sheetPath = (f: { driveId: string; itemId: string }, name: string) => `${book(f)}/worksheets/${encodeURIComponent(sheetName(name))}`;

/**
 * Writes the research's workbook to OneDrive and lets Excel compute it: pivots refreshed (their values
 * come from the models, with the user's rights), charts added on their sheets. Returns the file.
 */
export async function publishResearch(r: ResearchEntry, progress?: (step: string) => void) {
  progress?.("Writing the workbook");
  const bytes = await buildWorkbook(r.sheets, await tenantId());
  const base64 = await toBase64(new Blob([bytes as BlobPart]));
  let file = r.file;
  // The research's workbook is the app's: written whole each time.
  if (file) await writeItem({ data: { driveId: file.driveId, itemId: file.itemId, base64 } });
  else {
    // A research without workbook (made before workbooks were chosen by link): one in OneDrive.
    const drive = await graph("/me/drive?$select=id");
    const root = await graph("/me/drive/root?$select=id");
    const name = `${r.name.replace(/[^\p{L}\p{N} _-]+/gu, " ").trim() || "Research"}.xlsx`;
    const item = await uploadFile({ data: { driveId: drive.id, parentId: root.id, path: `${FOLDER}/${name}`, base64, type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", replace: true } });
    file = { driveId: drive.id as string, itemId: item.id, webUrl: item.webUrl };
  }
  const errors: string[] = [];
  for (const s of r.sheets.filter((x) => x.kind === "pivot")) {
    progress?.(`Refreshing "${s.name}" from the model`);
    await graph(`${sheetPath(file, s.name)}/pivotTables/refreshAll`, "POST", {}).catch((e) => errors.push(`${s.name}: ${e instanceof Error ? e.message : e}`));
  }
  for (const s of r.sheets) {
    if (s.kind !== "chart") continue;
    progress?.(`Drawing "${s.name}"`);
    try {
      const { address } = await graph(`${sheetPath(file, s.source)}/usedRange(valuesOnly=true)?$select=address`);
      const chart = await graph(`${sheetPath(file, s.name)}/charts/add`, "POST", { type: s.type, sourceData: address, seriesBy: "Auto" });
      const path = `${sheetPath(file, s.name)}/charts/${encodeURIComponent(chart.name)}`;
      await graph(path, "PATCH", { top: 10, left: 10, width: 720, height: 400 });
      await graph(`${path}/title`, "PATCH", { text: s.name, visible: true }).catch(() => undefined);
    } catch (e) {
      errors.push(`${s.name}: ${e instanceof Error ? e.message : e}`);
    }
  }
  await patch("researches", r.id, { file, updatedAt: Date.now() });
  return { file, errors };
}

/** Values of a sheet as Excel computed them (first `limit` rows). */
export async function sheetValues(r: ResearchEntry, name: string, limit = 30) {
  if (!r.file) throw new Error("The workbook is not published yet.");
  const { values, address } = await graph(`${sheetPath(r.file, name)}/usedRange(valuesOnly=true)?$select=values,address`);
  return { address: address as string, values: (values as unknown[][]).slice(0, limit + 1), total: (values as unknown[][]).length };
}

/** The workbook of a new research, from the link of an Excel file (SharePoint / OneDrive). */
export async function researchFile(link: string) {
  const f = await resolveExcel(link);
  return { name: f.name, file: { driveId: f.driveId, itemId: f.itemId, webUrl: f.webUrl } };
}
