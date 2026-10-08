/**
 * Semantic models as TMDL through the Fabric API. The model online is the source of truth for the
 * structure (types, formats, measures); the Excel file only holds the data.
 */
import type { Table } from "./excel";
import { fabric } from "./ms";
import type { FieldType } from "./schema";
import type { ExcelEntry } from "./store";

export type ModelFile = { path: string; text: string };

/** Fabric needs a workspace id: "My workspace" has one too. */
export async function workspaceId(groupId?: string) {
  if (groupId) return groupId;
  const { value } = await fabric("/workspaces");
  return value.find((w: { type: string }) => w.type === "Personal").id as string;
}

const q = (name: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
/** Field type → TMDL dataType and Power Query type (options, rich text and image links are text). */
const TYPES: Record<FieldType, [string, string]> = {
  string: ["string", "type text"],
  text: ["string", "type text"],
  option: ["string", "type text"],
  image: ["string", "type text"],
  number: ["double", "type number"],
  integer: ["int64", "Int64.Type"],
  boolean: ["boolean", "type logical"],
  date: ["dateTime", "type date"],
};

/** Table of an Excel file as TMDL, keeping the measures (and column settings) already written online. */
export function tableTmdl(e: ExcelEntry, table: Table, previous?: string) {
  const kept = (kind: "measure" | "column", name: string) =>
    previous?.match(new RegExp(`\\n(?:\\t///[^\\n]*\\n)*\\t${kind} ${escape(q(name))}\\b[\\s\\S]*?(?=\\n(?:\\t///[^\\n]*\\n)*\\t(?:column|measure|partition|annotation) |$)`))?.[0].slice(1);
  const columns = table.fields.map((f) => {
    const old = kept("column", f.name);
    // Settings written online are kept, unless the type changed.
    if (old?.includes(`dataType: ${TYPES[f.type][0]}`)) return old.trimEnd();
    return [
      f.description && `\t/// ${f.description.replace(/\n/g, " ")}`,
      `\tcolumn ${q(f.name)}`,
      `\t\tdataType: ${TYPES[f.type][0]}`,
      f.type === "date" && "\t\tformatString: Short Date",
      `\t\tsummarizeBy: ${f.type === "number" || f.type === "integer" ? "sum" : "none"}`,
      `\t\tsourceColumn: ${f.name}`,
      f.type === "image" && "\t\tdataCategory: ImageUrl",
    ]
      .filter(Boolean)
      .join("\n");
  });
  const measures = [...(previous?.matchAll(/\n(?:\t\/\/\/[^\n]*\n)*\tmeasure [\s\S]*?(?=\n(?:\t\/\/\/[^\n]*\n)*\t(?:column|measure|partition|annotation) |$)/g) ?? [])].map((m) => m[0].slice(1).trimEnd());
  const types = table.fields.map((f) => `{"${f.name}", ${TYPES[f.type][1]}}`).join(", ");
  return `table ${q(e.table)}

${[...columns, ...(measures.length ? measures : [`\tmeasure 'Row count' = COUNTROWS(${q(e.table)})`])].join("\n\n")}

\tpartition ${q(e.table)} = m
\t\tmode: import
\t\tsource =
\t\t\t\tlet
\t\t\t\t    Source = Excel.Workbook(Web.Contents("${e.fileUrl}"), null, true),
\t\t\t\t    Data = Source{[Item="${e.table}",Kind="Table"]}[Data],
\t\t\t\t    Typed = Table.TransformColumnTypes(Data, {${types}})
\t\t\t\tin
\t\t\t\t    Typed
`;
}
const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const encode = (files: ModelFile[]) => files.map((f) => ({ path: f.path, payload: btoa(String.fromCharCode(...new TextEncoder().encode(f.text))), payloadType: "InlineBase64" }));
const decode = (payload: string) => new TextDecoder().decode(Uint8Array.from(atob(payload), (c) => c.charCodeAt(0)));

/** New model with one table read from the Excel file. */
export async function createModel(e: ExcelEntry, table: Table, name: string) {
  const files: ModelFile[] = [
    { path: "definition.pbism", text: JSON.stringify({ version: "4.2", settings: {} }) },
    { path: "definition/database.tmdl", text: "database\n\tcompatibilityLevel: 1567\n" },
    { path: "definition/model.tmdl", text: `model Model\n\tculture: en-US\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n\tsourceQueryCulture: en-US\n\nref table ${q(e.table)}\n` },
    { path: `definition/tables/${e.table}.tmdl`, text: tableTmdl(e, table) },
  ];
  const ws = await workspaceId();
  const created = await fabric(`/workspaces/${ws}/semanticModels`, "POST", { displayName: name, definition: { parts: encode(files) } });
  if (created?.id) return created.id as string;
  const { value } = await fabric(`/workspaces/${ws}/semanticModels`);
  return value.find((m: { displayName: string }) => m.displayName === name).id as string;
}

export async function readModelFiles(datasetId: string, groupId?: string): Promise<ModelFile[]> {
  const result = await fabric(`/workspaces/${await workspaceId(groupId)}/semanticModels/${datasetId}/getDefinition?format=TMDL`, "POST");
  return result.definition.parts.filter((p: { path: string }) => p.path !== ".platform").map((p: { path: string; payload: string }) => ({ path: p.path, text: decode(p.payload) }));
}

export async function writeModelFiles(datasetId: string, files: ModelFile[], groupId?: string) {
  await fabric(`/workspaces/${await workspaceId(groupId)}/semanticModels/${datasetId}/updateDefinition`, "POST", { definition: { parts: encode(files) } });
}
