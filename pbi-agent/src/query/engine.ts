/**
 * One way to query any dataset of the database: remote semantic models in DAX
 * (executeQueries), local Excel models in SQL. Pages and the agent only use this.
 */
import { db } from "@/db/db";
import type { Row } from "@/db/schema";
import type { QueryResult } from "@/pbi/api";
import { executeDax } from "@/pbi/api";
import type { Connection, ModelContent } from "@/pbi/types";
import { executeSql } from "./local";

export type Source = {
  id: string;
  title: string;
  kind: "remote" | "excel";
  language: "DAX" | "SQL";
  workspaceId?: string;
  model: ModelContent;
  context?: string;
};

export async function loadSource(row: Row): Promise<Source> {
  const model = await db.readJson<ModelContent>(row.content_json);
  if (!model) throw new Error(`No model stored for ${row.title}: refresh it in the database.`);
  const connection = await db.readJson<Connection>(row.config_json);
  const kind = row.source === "excel" || connection?.kind === "excel" ? "excel" : "remote";
  return {
    id: row.id,
    title: String(row.title),
    kind,
    language: kind === "excel" ? "SQL" : "DAX",
    workspaceId: (row.workspace_id as string) || undefined,
    model,
    context: (row.context as string) || undefined,
  };
}

export const MAX_ROWS = 50_000;

export async function runQuery(source: Source, query: string): Promise<QueryResult> {
  const result =
    source.kind === "excel" ? await executeSql(source.id, query) : await executeDax({ datasetId: source.id, groupId: source.workspaceId }, query);
  if (result.rows.length > MAX_ROWS) result.rows = result.rows.slice(0, MAX_ROWS);
  return result;
}

/** Example query shown in an empty editor. */
export function sampleQuery(source: Source) {
  const table = source.model.tables.find((t) => t.columns.length) ?? source.model.tables[0];
  if (!table) return "";
  return source.language === "SQL"
    ? `SELECT TOP 100 * FROM [${table.name}]`
    : `EVALUATE\nTOPN(100, '${table.name}')`;
}

/** Results as CSV (for downloads and the agent). */
export function toCsv(result: QueryResult, limit = Infinity) {
  const cell = (v: unknown) => {
    const text = v === null || v === undefined ? "" : v instanceof Date ? v.toISOString() : String(v);
    return /[",\n;]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  };
  return [result.columns.map(cell).join(","), ...result.rows.slice(0, limit).map((r) => result.columns.map((c) => cell(r[c])).join(","))].join("\n");
}
