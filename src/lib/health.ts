/**
 * Health of the library's items, at a glance: what works and what will not (rights, last refresh,
 * readable file…), so a failing feature is explained before it is tried. Read-only checks.
 */
import { readTable } from "./excel";
import { pbi } from "./ms";
import { dax, scope } from "./pbi";
import type { DatasetEntry, ExcelEntry, ReportEntry } from "./store";

export type Check = { level: "ok" | "warn" | "error"; text: string };
const ok = (text: string): Check => ({ level: "ok", text });
const warn = (text: string): Check => ({ level: "warn", text });
const bad = (text: string): Check => ({ level: "error", text });
const why = (e: unknown) => (e instanceof Error ? e.message : String(e)).slice(0, 120);

const ago = (iso: string) => {
  const h = (Date.now() - new Date(iso).getTime()) / 3_600_000;
  return h < 1 ? `${Math.max(1, Math.round(h * 60))} min ago` : h < 48 ? `${Math.round(h)} h ago` : `${Math.round(h / 24)} days ago`;
};

/** Last refresh of a model: when, and whether it worked (needs access to its workspace). */
async function lastRefresh(id: string, groupId?: string): Promise<Check> {
  const last = (await pbi(`${scope(groupId)}/datasets/${id}/refreshes?$top=1`).catch(() => null))?.value?.[0];
  if (!last) return warn("Last refresh: unknown (no access to its workspace, or never refreshed)");
  if (last.status === "Failed") return bad(`Last refresh FAILED ${ago(last.startTime)}`);
  const days = (Date.now() - new Date(last.endTime ?? last.startTime).getTime()) / 86_400_000;
  return (days > 7 ? warn : ok)(`Data refreshed ${ago(last.endTime ?? last.startTime)}${last.status === "Unknown" ? " (running)" : ""}`);
}

export async function reportHealth(r: ReportEntry, model?: DatasetEntry): Promise<Check[]> {
  const out: Check[] = [];
  const found = await pbi(r.appId ? `/apps/${r.appId}/reports/${r.id}` : `${scope(r.groupId)}/reports/${r.id}`).then(() => true, (e) => (out.push(bad(`Not reachable: ${why(e)}`)), false));
  if (!found) return out;
  out.push(r.appId ? warn("Report of an app: view only, no definition (no guides links per visual from the builder, no editing)") : r.editable ? ok("Editable (builder, clone)") : ok("View rights"));
  if (r.datasetError || !model) out.push(warn(`Model not in the library${r.datasetError ? `: ${r.datasetError}` : ""} (the AI cannot run DAX)`));
  else out.push(await lastRefresh(model.id, model.groupId));
  if (!r.infoExcelId) out.push(warn("No info links file"));
  if (!r.guidesExcelId) out.push(warn("No guides file"));
  return out;
}

export async function datasetHealth(d: DatasetEntry): Promise<Check[]> {
  const out: Check[] = [];
  out.push(await dax(d, 'EVALUATE ROW("ok", 1)').then(() => ok("DAX queries allowed (Build right)"), (e) => bad(`DAX refused: ${why(e)} (the AI and the DAX / TCD builders cannot read it)`)));
  out.push(!d.info ? warn("Structure not read: use Read again") : d.info.source === "tmdl" ? ok("Full structure (write access)") : warn("Partial structure (no write access: no DAX of the measures, no parameters)"));
  if (d.info?.warnings.length) out.push(warn(d.info.warnings.join("; ")));
  out.push(await lastRefresh(d.id, d.groupId));
  return out;
}

export async function excelHealth(e: ExcelEntry): Promise<Check[]> {
  const out: Check[] = [];
  const table = await readTable(e).catch((err) => (out.push(bad(`Not readable: ${why(err)}`)), undefined));
  if (!table) return out;
  out.push(ok(`${table.rows.length} rows, ${table.fields.length} fields`));
  out.push(table.hasSchema ? ok("Schema saved in the file") : warn("No schema yet: the field types are guessed (save them in Fields)"));
  if (e.form) out.push(ok(`Form linked: ${e.form.title || "untitled"} (${e.form.questions.length} questions)`));
  if (table.fields.some((f) => f.type === "image") && !e.imagesFolder) out.push(warn("Image fields but no images folder: uploads are off"));
  return out;
}
