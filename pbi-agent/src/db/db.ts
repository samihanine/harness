/**
 * The local database: a folder holding reports.xlsx, datasets.xlsx and guides.xlsx.
 * Large JSON (report content, models) is stored next to them in json/<table>/<id>.<kind>.json,
 * the Excel cell keeping the relative path (an Excel cell holds at most 32,767 characters).
 */
import { useSyncExternalStore } from "react";
import { createStore, get, set } from "idb-keyval";
import { ExcelFile } from "./excel";
import type { Row, Schema } from "./schema";

const s = (name: string, extra: Partial<Schema["fields"][number]> = {}) => ({ name, type: "string" as const, ...extra });

export const SCHEMAS = {
  reports: {
    id: "reports",
    name: "Reports",
    sheet: "Reports",
    fields: [
      s("title", { required: true }),
      s("dataset_id"),
      s("workspace_id"),
      s("config_json"),
      s("content_json"),
      s("context", { type: "text" }),
      s("updated_at"),
    ],
  },
  datasets: {
    id: "datasets",
    name: "Datasets",
    sheet: "Datasets",
    fields: [
      s("title", { required: true }),
      { name: "source", type: "option" as const, options: [{ value: "remote", color: "blue" as const }, { value: "excel", color: "green" as const }] },
      s("workspace_id"),
      s("config_json"),
      s("content_json"),
      s("context", { type: "text" }),
      s("updated_at"),
    ],
  },
  guides: {
    id: "guides",
    name: "Guides",
    sheet: "Guides",
    fields: [s("title", { required: true }), s("content", { type: "text" }), s("visual_ids", { type: "text" })],
  },
} satisfies Record<string, Schema>;

export type TableName = keyof typeof SCHEMAS;
export const TABLE_NAMES = Object.keys(SCHEMAS) as TableName[];

type State =
  | { status: "loading" }
  | { status: "no-folder" }
  | { status: "permission"; folder: string }
  | { status: "error"; error: string }
  | { status: "ready"; folder: string };

const handles = createStore("pbi-agent", "handles");
const listeners = new Set<() => void>();
let version = 0;
const emit = () => {
  version++;
  listeners.forEach((l) => l());
};

class Database {
  state: State = { status: "loading" };
  folder?: FileSystemDirectoryHandle;
  tables = {} as Record<TableName, ExcelFile>;

  /** Opens the saved folder (asking for access needs a click: ask = true). */
  async open(ask = false) {
    try {
      const folder = await get<FileSystemDirectoryHandle>("folder", handles);
      if (!folder) return this.set({ status: "no-folder" });
      const mode = { mode: "readwrite" as const };
      if ((await folder.queryPermission(mode)) !== "granted" && !(ask && (await folder.requestPermission(mode)) === "granted"))
        return this.set({ status: "permission", folder: folder.name });
      this.folder = folder;
      for (const name of TABLE_NAMES) {
        const file = await folder.getFileHandle(`${name}.xlsx`, { create: true });
        this.tables[name] = await ExcelFile.open(file, SCHEMAS[name]);
        this.tables[name].subscribe(emit);
      }
      this.set({ status: "ready", folder: folder.name });
    } catch (error) {
      this.set({ status: "error", error: error instanceof Error ? error.message : String(error) });
    }
  }

  async choose(folder: FileSystemDirectoryHandle) {
    await set("folder", folder, handles);
    await this.open(true);
  }

  private set(state: State) {
    this.state = state;
    emit();
  }

  rows(table: TableName): Row[] {
    return this.state.status === "ready" ? this.tables[table].rows : [];
  }

  row(table: TableName, id: string) {
    return this.rows(table).find((r) => r.id === id);
  }

  /** Inserts or updates a row by id. */
  upsert(table: TableName, id: string, values: Record<string, unknown>) {
    const excel = this.tables[table];
    if (excel.rows.some((r) => r.id === id)) excel.update([{ id, values }]);
    else excel.insert([{ ...values, id }]);
  }

  async writeJson(table: TableName, id: string, kind: string, data: unknown) {
    const dir = await (await this.folder!.getDirectoryHandle("json", { create: true })).getDirectoryHandle(table, { create: true });
    const name = `${id}.${kind}.json`;
    const stream = await (await dir.getFileHandle(name, { create: true })).createWritable();
    await stream.write(JSON.stringify(data, null, 1));
    await stream.close();
    jsonCache.delete(`json/${table}/${name}`);
    return `json/${table}/${name}`;
  }

  /** Reads a JSON file referenced by a cell (or the cell itself when it holds JSON). */
  async readJson<T>(ref: unknown): Promise<T | undefined> {
    const value = String(ref ?? "").trim();
    if (!value) return undefined;
    if (value.startsWith("{") || value.startsWith("[")) return JSON.parse(value) as T;
    if (!jsonCache.has(value))
      jsonCache.set(
        value,
        (async () => {
          let dir = this.folder!;
          const parts = value.split("/");
          for (const part of parts.slice(0, -1)) dir = await dir.getDirectoryHandle(part);
          return JSON.parse(await (await (await dir.getFileHandle(parts.at(-1)!)).getFile()).text());
        })(),
      );
    return jsonCache.get(value) as Promise<T>;
  }

  async removeJson(table: TableName, id: string) {
    const dir = await (await this.folder!.getDirectoryHandle("json", { create: true })).getDirectoryHandle(table, { create: true });
    for await (const name of dir.keys())
      if (name.startsWith(`${id}.`)) {
        await dir.removeEntry(name);
        jsonCache.delete(`json/${table}/${name}`);
      }
  }

  async remove(table: TableName, id: string) {
    this.tables[table].remove([id]);
    await this.removeJson(table, id).catch(() => undefined);
  }
}

const jsonCache = new Map<string, Promise<unknown>>();

export const db = new Database();
void db.open();

/** Re-renders on any database change; returns the database. */
export function useDb() {
  useSyncExternalStore(
    (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    () => version,
  );
  return db;
}

/** Follows changes made to the Excel files outside the app. */
setInterval(() => {
  if (db.state.status === "ready") for (const name of TABLE_NAMES) void db.tables[name].reloadIfChanged();
}, 5_000);
