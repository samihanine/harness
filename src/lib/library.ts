/**
 * Library entries: what is read from Power BI / SharePoint / Forms when a link is added or read again,
 * and what the user may do with each (access). Kept in this browser (IndexedDB); every entry can be
 * downloaded as JSON and imported again.
 */
import { applySchemaToExcel, createTable, detectFields, resolveExcel, type Progress } from "./excel";
import { fieldsFromQuestions, layoutFromQuestions, readFormRef, withFormTypes } from "./forms";
import { fabric, pbi } from "./ms";
import { addSamples, readModel } from "./model";
import { parseDatasetUrl, parseReportUrl, scope, type Ref } from "./pbi";
import { definition, legacyDefinition } from "./report-def";
import { COLORS, ID_FIELD, fieldsErrors, isIdField, type Field, type FormSection } from "./schema";
import { all, find, patch, put, remove, type ContextEntry, type DatasetEntry, type ExcelRef, type Name, type ReportEntry, type SchemaEntry } from "./store";
import { workspaceId } from "./tmdl";

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
  const { createBlankReport } = await import("./research");
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

/**
 * A new semantic model (in My workspace) reading Excel tables of SharePoint / OneDrive: one model
 * table per Excel table, typed after its columns. Its data comes with its first refresh (from the
 * Dataset editor or the library).
 */
export async function createDatasetFromExcel(name: string, sources: { link: string; table?: string }[]) {
  const tables: { name: string; tmdl: string }[] = [];
  for (const s of sources) {
    const file = await resolveExcel(s.link);
    const table = s.table && file.tables.includes(s.table) ? s.table : file.tables[0];
    if (!table) throw new Error(`"${file.name}" has no Excel table: select the data in Excel and use Insert › Table.`);
    const excel: ExcelRef = { link: s.link, driveId: file.driveId, itemId: file.itemId, name: file.name, table, fileUrl: file.fileUrl };
    let tableName = table;
    for (let i = 2; tables.some((t) => t.name.toLowerCase() === tableName.toLowerCase()); i++) tableName = `${table} ${i}`;
    tables.push({ name: tableName, tmdl: excelTableTmdl(tableName, excel, await detectFields(excel)) });
  }
  return createModel(name, tables);
}

/** A new semantic model in My workspace from TMDL tables; added to the app. */
async function createModel(name: string, tables: { name: string; tmdl: string }[]) {
  const encode = (text: string) => btoa(String.fromCharCode(...new TextEncoder().encode(text)));
  const parts = [
    { path: "definition.pbism", text: JSON.stringify({ version: "4.2", settings: {} }) },
    { path: "definition/database.tmdl", text: "database\n\tcompatibilityLevel: 1567\n" },
    { path: "definition/model.tmdl", text: `model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n\tsourceQueryCulture: en-US\n\n${tables.map((t) => `ref table ${q(t.name)}`).join("\n")}\n` },
    ...tables.map((t) => ({ path: `definition/tables/${t.name}.tmdl`, text: t.tmdl })),
  ].map((p) => ({ path: p.path, payload: encode(p.text), payloadType: "InlineBase64" }));
  const ws = await workspaceId();
  const created = await fabric(`/workspaces/${ws}/semanticModels`, "POST", { displayName: name, definition: { parts } });
  const id = (created?.id as string | undefined) ?? (await fabric(`/workspaces/${ws}/semanticModels`)).value.find((m: { displayName: string }) => m.displayName === name)?.id;
  if (!id) throw new Error("The model was created but could not be found in My workspace.");
  return addDataset(id, ws, name);
}

/** The cells of a CSV text (quotes handled), with its delimiter (the one splitting the header most). */
export function parseCsv(text: string) {
  const firstLine = text.slice(0, text.indexOf("\n") >>> 0 || undefined);
  const delimiter = [",", ";", "\t", "|"].sort((a, b) => firstLine.split(b).length - firstLine.split(a).length)[0];
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
    else if (c === delimiter) row.push(cell), (cell = "");
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(cell);
      if (row.some((x) => x !== "")) rows.push(row);
      (row = []), (cell = "");
    } else cell += c;
  }
  row.push(cell);
  if (row.some((x) => x !== "")) rows.push(row);
  return { delimiter, header: rows[0] ?? [], rows: rows.slice(1) };
}

/**
 * A new semantic model from a CSV file: the file is put in OneDrive (Hub/Datasets), and a model table
 * reads it (Power Query, typed from its values: numbers, dates, text). Its data comes with a refresh.
 */
export async function createDatasetFromCsv(name: string, file: File) {
  if (file.size > 50 * 1024 * 1024) throw new Error("CSV files up to 50 MB.");
  const text = await file.text();
  const { delimiter, header, rows } = parseCsv(text);
  if (!header.length) throw new Error("The CSV file has no header line.");
  const sample = rows.slice(0, 200);
  const typeOf = (i: number) => {
    const values = sample.map((r) => (r[i] ?? "").trim()).filter(Boolean);
    if (values.length && values.every((v) => /^-?\d+([.,]\d+)?$/.test(v))) return values.every((v) => /^-?\d+$/.test(v)) ? ["int64", "Int64.Type"] : ["double", "type number"];
    if (values.length && values.every((v) => /^\d{4}-\d{2}-\d{2}/.test(v))) return ["dateTime", "type date"];
    return ["string", "type text"];
  };
  const { uploadFile } = await import("@/server/ms");
  const { graph } = await import("./ms");
  const drive = await graph("/me/drive?$select=id,webUrl");
  const root = await graph("/me/drive/root?$select=id");
  const fileName = `${name.replace(/[^\p{L}\p{N} _-]+/gu, " ").trim() || "Data"}.csv`;
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  await uploadFile({ data: { driveId: drive.id, parentId: root.id, path: `Hub/Datasets/${fileName}`, base64: btoa(binary), type: "text/csv", replace: true } });
  const url = `${drive.webUrl}/Hub/Datasets/${encodeURI(fileName)}`;
  const tableName = name.trim() || "Data";
  const types = header.map((h, i) => ({ name: h.trim() || `Column ${i + 1}`, type: typeOf(i) }));
  const tmdl = `table ${q(tableName)}

${types.map((c) => [`\tcolumn ${q(c.name)}`, `\t\tdataType: ${c.type[0]}`, c.type[0] === "dateTime" && "\t\tformatString: Short Date", `\t\tsummarizeBy: ${c.type[0] === "string" || c.type[0] === "dateTime" ? "none" : "sum"}`, `\t\tsourceColumn: ${c.name}`].filter(Boolean).join("\n")).join("\n\n")}

\tmeasure 'Rows of ${tableName.replace(/'/g, "''")}' = COUNTROWS(${q(tableName)})

\tpartition ${q(tableName)} = m
\t\tmode: import
\t\tsource =
\t\t\t\tlet
\t\t\t\t    Source = Csv.Document(Web.Contents("${url}"), [Delimiter="${delimiter === "\t" ? "#(tab)" : delimiter}", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
\t\t\t\t    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
\t\t\t\t    Typed = Table.TransformColumnTypes(Promoted, {${types.map((c) => `{"${c.name.replace(/"/g, '""')}", ${c.type[1]}}`).join(", ")}}, "en-US")
\t\t\t\tin
\t\t\t\t    Typed
`;
  return createModel(name, [{ name: tableName, tmdl }]);
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
const BUILT_INS: SchemaEntry[] = [
  { id: "guides", builtIn: "guides", title: "Guides (report viewer)", fields: GUIDE_FIELDS, contextIds: [] },
];

/** The built-in schemas exist in every library, with their current fields. */
export async function ensureBuiltIns() {
  const schemas = await all("schemas");
  // The Links feature was removed: its built-in schema goes too (the Excel file is kept).
  if (schemas.some((s) => s.id === "links" && (s.builtIn as string) === "links")) await remove("schemas", "links");
  for (const b of BUILT_INS) {
    const old = schemas.find((s) => s.id === b.id);
    if (!old) await put("schemas", b);
    else if (JSON.stringify(old.fields) !== JSON.stringify(b.fields)) await patch("schemas", b.id, { fields: b.fields });
  }
}

/** An Excel table of a link: the given table, else the first one. */
async function excelOf(link: string, table?: string): Promise<ExcelRef & { tables: string[] }> {
  const file = await resolveExcel(link);
  const name = table && file.tables.includes(table) ? table : file.tables[0];
  return { link, driveId: file.driveId, itemId: file.itemId, name: file.name, table: name ?? "", fileUrl: file.fileUrl, webUrl: file.webUrl, tables: file.tables };
}

/**
 * Adds a schema from an Excel table (its columns are detected, typed from the data or its "_schema"
 * sheet) and, optionally, the Microsoft Form filling it (its questions give the types and choices).
 */
export async function addSchema({ excel: link, table, form: formLink, title }: { excel: string; table?: string; form?: string; title?: string }) {
  // Read first: a wrong form link fails the add before anything is kept.
  const form = formLink?.trim() ? await readFormRef(formLink) : undefined;
  const { tables, ...excel } = await excelOf(link, table);
  if (!excel.table) throw new Error(`"${excel.name}" has no Excel table: select the data in Excel and use Insert › Table.`);
  let fields = await detectFields(excel);
  let sections: FormSection[] = [];
  if (form) ({ fields, sections } = layoutFromQuestions(form.questions, withFormTypes(fields, form.questions)));
  const old = (await all("schemas")).find((s) => s.excel?.itemId === excel.itemId && s.excel.table === excel.table);
  return put("schemas", {
    ...old,
    id: old?.id ?? crypto.randomUUID(),
    title: title?.trim() || old?.title || (tables.length > 1 ? `${excel.name} › ${excel.table}` : excel.name),
    excel,
    form: form ? { url: form.url, id: form.id, title: form.title } : old?.form,
    fields: old?.fields ?? fields,
    sections: old?.sections ?? (sections.length ? sections : undefined),
    contextIds: old?.contextIds ?? [],
  });
}

/** Links a form to a schema; its questions give the types of the matching columns when `retype`. */
export async function linkForm(schema: SchemaEntry, link: string, retype = false) {
  const form = await readFormRef(link);
  // Retyped: the matching columns take the questions' types, and the form's sections and branches.
  const layout = retype ? layoutFromQuestions(form.questions, withFormTypes(schema.fields, form.questions)) : undefined;
  await patch("schemas", schema.id, { form: { url: form.url, id: form.id, title: form.title }, ...(layout && { fields: layout.fields, sections: layout.sections.length ? layout.sections : undefined }) });
  return form;
}

/** Fields of a form alone (one per question), to compare or start from. */
export const fieldsOfForm = async (link: string) => fieldsFromQuestions((await readFormRef(link)).questions);

/**
 * The Excel file of a built-in schema: a file without table gets one (header = the fields), and the
 * missing columns are added (with the "_schema" sheet and the styling).
 */
export async function setBuiltInExcel(schema: SchemaEntry, link: string, progress?: Progress) {
  let excel = await excelOf(link);
  if (!excel.table) {
    progress?.("Creating the table", 0, 1);
    await createTable(excel, schema.fields.map((f) => f.name));
    excel = await excelOf(link);
  }
  const { tables: _, ...ref } = excel;
  await applySchemaToExcel(ref, schema.fields, progress);
  await patch("schemas", schema.id, { excel: ref });
}

/** Saves a schema's fields (and its form's sections) after checking them (names, types, choices, branches). */
export async function saveFields(schema: SchemaEntry, fields: Field[], sections: FormSection[] | undefined = schema.sections) {
  if (schema.builtIn) throw new Error("The fields of a built-in schema cannot change.");
  const list = fields.some(isIdField) ? fields : [ID_FIELD, ...fields];
  const errors = fieldsErrors(list, sections);
  if (errors.length) throw new Error(errors.join(" "));
  await patch("schemas", schema.id, { fields: list, sections: sections?.length ? sections : undefined });
  return list;
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
