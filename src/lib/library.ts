/**
 * Library entries: what is read from Power BI / SharePoint when a link is added or read again,
 * and what the user may do with each (access). Kept in this browser (IndexedDB); every entry can be
 * downloaded as JSON and imported again.
 */
import { detectFields, resolveExcel } from "@/services/excel";
import { fabric, pbi } from "@/services/ms";
import { addSamples, readModel } from "@/services/model";
import { parseDatasetUrl, parseReportUrl, scope, type Ref } from "@/services/pbi";
import { addEmptyGuides } from "@/features/trainer/guides";
import { definition, legacyDefinition } from "@/services/report-def";
import { describeReport } from "@/services/report-schema";
import { COLORS, ID_FIELD, fieldsErrors, isIdField, type Field, type FormSection } from "@/features/excel/schema";
import { all, find, patch, put, remove, type ContextEntry, type DatasetEntry, type ExcelRef, type Name, type ReportEntry, type SchemaEntry } from "./store";
import { workspaceId } from "@/services/tmdl";

const message = (e: unknown) => (e instanceof Error ? e.message : String(e));

/* --------------------------------- Reports --------------------------------- */

/** What the user may do with a report: edit when its PBIR definition is readable (the editor needs it). */
async function reportAccess(ref: Ref): Promise<Pick<ReportEntry, "access" | "accessNote">> {
  if (ref.appId) return { access: "view", accessNote: "Report of a Power BI app: view only." };
  try {
    await definition(ref);
    return { access: "edit", accessNote: undefined };
  } catch (e) {
    // Readable in the older format: editable in Power BI, not by the editor's JSON tools.
    const legacy = await legacyDefinition(ref).then(() => true, () => false);
    return legacy
      ? { access: "view", accessNote: "Saved in the older PBIR-Legacy format: open it once in the Power BI editor and save it to convert it, then read it again." }
      : { access: "view", accessNote: `Definition not readable (${message(e).slice(0, 160)}): view only.` };
  }
}

export async function addReport(link: string): Promise<ReportEntry> {
  const ref = parseReportUrl(link);
  const r = await pbi(ref.appId ? `/apps/${ref.appId}/reports/${ref.id}` : `${scope(ref.groupId)}/reports/${ref.id}`);
  const old = await find("reports", ref.id);
  const entry = await put("reports", {
    ...old,
    id: ref.id,
    groupId: ref.groupId,
    appId: ref.appId,
    name: old?.renamed ? old.name : r.name,
    url: r.webUrl ?? link,
    datasetId: r.datasetId,
    datasetGroupId: r.datasetWorkspaceId ?? ref.groupId,
    contextIds: old?.contextIds ?? [],
    ...(await reportAccess(ref)),
  });
  // Its semantic model comes with it (in the background: reading a model takes a while).
  if (!(await find("datasets", r.datasetId))) void addDataset(r.datasetId, entry.datasetGroupId, `${r.name} (model)`).catch(() => undefined);
  // Its schema, then an empty guide for each visual (the guides file then lists them all), in the background too.
  void (old?.schema ? Promise.resolve(old.schema) : describeReport(entry)).then((schema) => addEmptyGuides(entry.id, schema)).catch(() => undefined);
  return entry;
}

/**
 * A new report on a semantic model (its link or id), live-connected to it, created in My workspace in
 * the format the Builder edits (PBIR), with one blank page; added to the app. Needs Build (or edit)
 * on the model.
 */
export async function createReportFromModel(modelLink: string, name?: string) {
  const model = await addDataset(modelLink);
  if (model.access === "view") throw new Error(`You cannot build on "${model.name}": a report on it needs the Build permission on the semantic model.${model.error ? ` (${model.error})` : ""}`);
  const title = name?.trim() || `${model.name} — report`;
  const { createBlankReport } = await import("@/features/builder/new-report");
  const id = await createBlankReport(title, { id: model.id, groupId: model.groupId });
  return addReport(`https://app.powerbi.com/groups/me/reports/${id}`);
}

/** Why a report cannot be edited by the Builder (a report of a Power BI app), with the way out. */
export const notEditable = (entry: ReportEntry) =>
  `"${entry.name}" is a report of a Power BI app: it cannot be edited. To build on the same data, create a new report from its semantic model (Builder › "New report from a semantic model…").`;

export async function refreshReport(entry: ReportEntry) {
  return addReport(entry.appId ? `https://app.powerbi.com/groups/me/apps/${entry.appId}/reports/${entry.id}` : `https://app.powerbi.com/groups/${entry.groupId ?? "me"}/reports/${entry.id}`);
}

/* -------------------------------- Datasets -------------------------------- */

/**
 * Adds a semantic model. Access: edit when its definition (TMDL) is readable, build when DAX queries
 * work (a model shared through an app), view otherwise (kept, with the reason). Value samples are
 * read in the background.
 */
export async function addDataset(linkOrId: string, groupId?: string, fallbackName?: string): Promise<DatasetEntry> {
  const ref = groupId ? { id: linkOrId, groupId } : parseDatasetUrl(linkOrId);
  const old = await find("datasets", ref.id);
  const d = await pbi(`${scope(ref.groupId)}/datasets/${ref.id}`).catch(() => null);
  const target: Ref = d ? ref : { id: ref.id };
  const base = { ...old, id: ref.id, groupId: target.groupId, name: old?.renamed ? old.name : (d?.name ?? old?.name ?? fallbackName ?? `Semantic model ${ref.id.slice(0, 8)}`), contextIds: old?.contextIds ?? [] };
  try {
    const info = await readModel(target, false);
    const entry = await put("datasets", { ...base, info, access: info.source === "tmdl" ? "edit" : "build", error: undefined });
    void addSamples(target, info).then(() => patch("datasets", ref.id, { info }), () => undefined);
    return entry;
  } catch (e) {
    return put("datasets", { ...base, access: "view", error: `Structure not readable: ${message(e)} (you need the Build permission on the model).` });
  }
}

export const refreshDataset = (entry: DatasetEntry) => addDataset(entry.id, entry.groupId, entry.name);

const q = (name: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
/** Field type → TMDL dataType and Power Query type (choices, long text and file links are text). */
const TYPES: Record<Field["type"], [string, string]> = {
  string: ["string", "type text"],
  text: ["string", "type text"],
  option: ["string", "type text"],
  url: ["string", "type text"],
  image: ["string", "type text"],
  file: ["string", "type text"],
  number: ["double", "type number"],
  rating: ["int64", "Int64.Type"],
  date: ["dateTime", "type date"],
};

/** An Excel table as a TMDL table: typed columns read from the file (Power Query), and a row count measure. */
function excelTableTmdl(name: string, excel: ExcelRef, fields: Field[]) {
  const columns = fields.map((f) =>
    [
      f.description && `\t/// ${f.description.replace(/\n/g, " ")}`,
      `\tcolumn ${q(f.name)}`,
      `\t\tdataType: ${TYPES[f.type][0]}`,
      f.type === "date" && "\t\tformatString: Short Date",
      `\t\tsummarizeBy: ${f.type === "number" && !isIdField(f) ? "sum" : "none"}`,
      `\t\tsourceColumn: ${f.name}`,
      (f.type === "image" || f.type === "file") && "\t\tdataCategory: ImageUrl",
      f.type === "url" && "\t\tdataCategory: WebUrl",
    ]
      .filter(Boolean)
      .join("\n"),
  );
  const types = fields.map((f) => `{"${f.name.replace(/"/g, '""')}", ${TYPES[f.type][1]}}`).join(", ");
  return `table ${q(name)}

${columns.join("\n\n")}

\tmeasure 'Rows of ${name.replace(/'/g, "''")}' = COUNTROWS(${q(name)})

\tpartition ${q(name)} = m
\t\tmode: import
\t\tsource =
\t\t\t\tlet
\t\t\t\t    Source = Excel.Workbook(Web.Contents("${excel.fileUrl}"), null, true),
\t\t\t\t    Data = Source{[Item="${excel.table}",Kind="Table"]}[Data],
\t\t\t\t    Typed = Table.TransformColumnTypes(Data, {${types}})
\t\t\t\tin
\t\t\t\t    Typed
`;
}

/* --------------------------------- Schemas --------------------------------- */

/** Built-in schemas of the report viewer: fixed fields; one Excel file holds the rows of every report (report_id). */
export const GUIDE_FIELDS: Field[] = [
  ID_FIELD,
  { name: "report_id", type: "string", required: true, description: "Id of the Power BI report" },
  { name: "page", type: "string", required: true, description: "Page name (id) of the report" },
  { name: "visual", type: "string", required: true, description: "Visual name (id) on that page" },
  { name: "title", type: "string", required: true },
  { name: "content", type: "text", description: "The guide, in markdown" },
];
/**
 * The guides now live in this browser (src/features/trainer/guides.ts): the built-in schemas of older versions (guides
 * and links in Excel files) leave the library (their Excel files are kept).
 */
export async function ensureBuiltIns() {
  for (const s of await all("schemas")) if (s.builtIn || s.id === "links") await remove("schemas", s.id);
}

/** An Excel table of a link: the given table, else the first one. */
async function excelOf(link: string, table?: string): Promise<ExcelRef & { tables: string[] }> {
  const file = await resolveExcel(link);
  const name = table && file.tables.includes(table) ? table : file.tables[0];
  return { link, driveId: file.driveId, itemId: file.itemId, name: file.name, table: name ?? "", fileUrl: file.fileUrl, webUrl: file.webUrl, tables: file.tables };
}

/** Adds an Excel table (its columns are the header of the table, read again each time it is opened). */
export async function addSchema({ excel: link, table, title }: { excel: string; table?: string; title?: string }) {
  const { tables, ...excel } = await excelOf(link, table);
  if (!excel.table) throw new Error(`"${excel.name}" has no Excel table: select the data in Excel and use Insert › Table.`);
  const fields = await detectFields(excel);
  const old = (await all("schemas")).find((s) => s.excel?.itemId === excel.itemId && s.excel.table === excel.table);
  return put("schemas", {
    ...old,
    id: old?.id ?? crypto.randomUUID(),
    title: title?.trim() || old?.title || (tables.length > 1 ? `${excel.name} › ${excel.table}` : excel.name),
    excel,
    fields,
    contextIds: old?.contextIds ?? [],
  });
}

/* -------------------------------- Contexts -------------------------------- */

/** Markdown files from the computer (one context each; a file with the same name replaces it). */
export async function addContexts(files: FileList | File[]) {
  const existing = await all("contexts");
  const added: ContextEntry[] = [];
  for (const file of Array.from(files)) {
    if (file.size > 500_000) throw new Error(`${file.name}: a context is a markdown file up to 500 KB.`);
    const name = file.name.replace(/\.(md|markdown|txt)$/i, "");
    const old = existing.find((c) => c.name === name);
    added.push(await put("contexts", { id: old?.id ?? crypto.randomUUID(), name, text: await file.text(), updatedAt: Date.now() }));
  }
  return added;
}

/** Contexts linked to an entry, as one text for the AI (each with its name). Long contexts are cut. */
export async function contextsText(ids: string[] = [], max = 24_000) {
  const contexts = (await all("contexts")).filter((c) => ids.includes(c.id));
  let left = max;
  return contexts
    .map((c) => {
      const text = c.text.length > left ? `${c.text.slice(0, Math.max(0, left))}\n…[cut: the contexts are too long]` : c.text;
      left -= text.length;
      return `### Context "${c.name}"\n${text}`;
    })
    .join("\n\n");
}

/* ------------------------------ Download / import ------------------------------ */

export type Kind = "reports" | "datasets" | "schemas" | "contexts" | "researches" | "decks";
const KINDS: Kind[] = ["reports", "datasets", "schemas", "contexts", "researches", "decks"];

/** Saves a file in the browser. */
export function downloadText(name: string, text: string, type = "application/json") {
  const url = URL.createObjectURL(new Blob([text], { type }));
  Object.assign(document.createElement("a"), { href: url, download: name.replace(/[^\p{L}\p{N}._ -]+/gu, "-") }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** One entry as a JSON file (a context also as its markdown file). */
export function downloadEntry(kind: Kind, entry: { id: string; name?: string; title?: string }) {
  if (kind === "contexts") return downloadText(`${(entry as ContextEntry).name}.md`, (entry as ContextEntry).text, "text/markdown");
  downloadText(`${entry.name ?? entry.title ?? entry.id}.${kind.slice(0, -1)}.json`, JSON.stringify({ hub: kind, entry }, null, 2));
}

/** Entries downloaded before (one per file, or a whole collection), put back in the library. */
export async function importEntries(files: FileList | File[]) {
  let count = 0;
  for (const file of Array.from(files)) {
    if (/\.(md|markdown|txt)$/i.test(file.name)) {
      count += (await addContexts([file])).length;
      continue;
    }
    const data = JSON.parse(await file.text());
    const items: { hub: Kind; entry: { id: string } }[] = Array.isArray(data) ? data : [data];
    for (const item of items) {
      if (!KINDS.includes(item?.hub) || typeof item.entry?.id !== "string") throw new Error(`${file.name}: not an entry downloaded from the library.`);
      await put(item.hub as Name, { contextIds: [], ...item.entry } as never);
      count++;
    }
  }
  return count;
}
