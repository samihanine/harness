/**
 * An Excel table kept in step with its file. The content is live: each change is written at once
 * (Graph $batch), on rows found again in a fresh read of the file (a row moved or changed in Excel
 * meanwhile is never overwritten by position). The structure (columns, their order and names, their
 * types: drop-down lists and colors) and the look are written by one sync pass, run after each column
 * change. The file's version is watched: a change made in Excel is read again.
 */
import { useCallback, useEffect, useRef, useState } from "react";
import { patch } from "@/lib/store";
import { errorText } from "@/services/ms";
import { addColumns, appendRows, deleteColumns, deleteRows, formatTable, readAppData, readTable, rewriteTable, versionOf, writeAppData, writeRows, type Cell, type ColumnLook } from "@/services/excel";
import { checkValue, columnOf, fileName, readData, reconcile, rowCells, type AppData, type Column, type Row, type TableData, type TableEntry, type View } from "./table";
import { writeChoicesWhenSaved } from "./xlsx";

export type SyncStatus = { state: "reading" | "ready" | "writing" | "syncing" | "error"; message?: string; at?: number; by?: string };

const LOOKS: Record<Column["type"], ColumnLook> = {
  text: { width: 170, wrap: true },
  longtext: { width: 330, wrap: true },
  date: { width: 90, wrap: false, numberFormat: "yyyy-mm-dd" },
  option: { width: 120, wrap: false },
};

const same = (a: Cell[] | undefined, b: Cell[] | undefined) => JSON.stringify(a ?? []) === JSON.stringify(b ?? []);

/** Where a row seen by the app is now in the file (same position if unchanged, else the first identical row). */
function locate(fresh: Cell[][], seen: Cell[] | undefined, index: number) {
  if (seen && same(fresh[index], seen)) return index;
  const found = seen ? fresh.findIndex((r) => same(r, seen)) : -1;
  if (found < 0) throw new Error(`Row ${index + 1} was changed in Excel meanwhile: the table was read again, try again.`);
  return found;
}

export type ExcelTable = ReturnType<typeof useExcelTable>;

export function useExcelTable(entry: TableEntry | undefined) {
  const [data, setData] = useState<TableData>();
  const [status, setStatus] = useState<SyncStatus>({ state: "reading" });
  const [dirty, setDirty] = useState(false);
  const live = useRef({ entry, data });
  live.current = { entry, data };
  const queue = useRef<Promise<unknown>>(Promise.resolve());
  const busy = useRef(0);
  const version = useRef<string | undefined>(undefined);
  const key = entry ? `${entry.id}|${entry.excel.itemId}|${entry.excel.table}` : undefined;

  /** Runs after the writes queued before it (row positions stay right); the status says what happens. */
  const queued = useCallback(<T,>(state: SyncStatus["state"], task: () => Promise<T>) => {
    const run = queue.current.then(async () => {
      busy.current++;
      setStatus((s) => ({ ...s, state }));
      try {
        const out = await task();
        setStatus({ state: "ready", at: Date.now() });
        return out;
      } catch (e) {
        setStatus({ state: "error", message: errorText(e), at: Date.now() });
        throw e;
      } finally {
        busy.current--;
        void versionOf(live.current.entry!.excel).then((v) => (version.current = v.eTag), () => undefined);
      }
    });
    queue.current = run.catch(() => undefined);
    return run;
  }, []);

  /** Reads the file: its columns against the app's (the file's copy of the types and views when newer), its rows. */
  const read = useCallback(async () => {
    const e = live.current.entry;
    if (!e) return undefined;
    const { header, rows } = await readTable(e.excel);
    let columns = e.columns;
    let views = e.views;
    const kept = await readAppData<AppData>(e.excel);
    if (kept && (!e.columns.length || kept.at > (e.changedAt ?? 0))) ({ columns, views } = kept);
    const r = reconcile(columns, header, rows, e.removed);
    if (JSON.stringify(r.columns) !== JSON.stringify(e.columns) || views !== e.views) await patch("tables", e.id, { columns: r.columns, views: views.length ? views : e.views });
    const next = readData(r.columns, header, rows);
    if (live.current.entry?.id === e.id) {
      setData(next);
      setDirty(r.dirty);
    }
    return { ...next, columns: r.columns };
  }, []);

  const reload = useCallback(() => queued("reading", read), [queued, read]);

  useEffect(() => {
    setData(undefined);
    setDirty(false);
    if (!key) return;
    void reload().catch(() => undefined);
    // A change saved in Excel (by anyone) is read again; not while the app writes.
    const timer = setInterval(async () => {
      const e = live.current.entry;
      if (!e || busy.current || document.hidden) return;
      const v = await versionOf(e.excel).catch(() => undefined);
      if (v && version.current && v.eTag !== version.current) {
        version.current = v.eTag;
        await reload().catch(() => undefined);
        setStatus({ state: "ready", at: Date.now(), message: `Updated from Excel${v.by ? ` (${v.by})` : ""}.` });
      } else if (v) version.current = v.eTag;
    }, 20_000);
    return () => clearInterval(timer);
  }, [key, reload]);

  const entryOf = () => {
    const e = live.current.entry;
    if (!e) throw new Error("No Excel table is chosen.");
    return e;
  };
  /** Values checked against their columns (an unknown column or option is an error the AI can fix). */
  const checked = (values: Record<string, unknown>) => {
    const columns = entryOf().columns;
    return Object.fromEntries(
      Object.entries(values).map(([name, v]) => {
        const c = columnOf(columns, name);
        if (!c) throw new Error(`No column "${name}": ${columns.map((x) => `"${x.name}"`).join(", ")}.`);
        return [c.name, checkValue(c, v)];
      }),
    ) as Row;
  };

  /* ---------------------------------- Rows ---------------------------------- */

  /** Changes rows (by their position in the data the app shows): each found again in the file first. */
  const update = (changes: { index: number; values: Record<string, unknown> }[]) =>
    queued("writing", async () => {
      const e = entryOf();
      const seen = live.current.data;
      const fresh = await readTable(e.excel);
      const writes = changes.map(({ index, values }) => {
        const at = locate(fresh.rows, seen?.cells[index], index);
        return { index: at, cells: rowCells(e.columns, fresh.header, fresh.rows[at], checked(values)) };
      });
      await writeRows(e.excel, writes);
      for (const w of writes) fresh.rows[w.index] = w.cells;
      setData(readData(e.columns, fresh.header, fresh.rows));
      return writes.length;
    });

  /** Adds rows at the end; returns their positions. */
  const add = (rows: Record<string, unknown>[]) =>
    queued("writing", async () => {
      const e = entryOf();
      const fresh = await readTable(e.excel);
      // A new table has one blank row: filled first instead of left empty.
      const blank = fresh.rows.length === 1 && fresh.rows[0].every((c) => c === "" || c === null) ? 0 : -1;
      const cells = rows.map((r) => rowCells(e.columns, fresh.header, undefined, checked(r)));
      if (blank === 0 && cells.length) {
        await writeRows(e.excel, [{ index: 0, cells: cells[0] }]);
        await appendRows(e.excel, cells.slice(1));
      } else await appendRows(e.excel, cells);
      const all = blank === 0 ? cells : [...fresh.rows, ...cells];
      setData(readData(e.columns, fresh.header, all));
      return cells.map((_, i) => all.length - cells.length + i);
    });

  /** Deletes rows (by their position in the data shown), each found again in the file first. */
  const remove = (indexes: number[]) =>
    queued("writing", async () => {
      const e = entryOf();
      const seen = live.current.data;
      const fresh = await readTable(e.excel);
      const at = indexes.map((i) => locate(fresh.rows, seen?.cells[i], i));
      await deleteRows(e.excel, at);
      const left = fresh.rows.filter((_, i) => !at.includes(i));
      setData(readData(e.columns, fresh.header, left));
      return at.length;
    });

  /* --------------------------------- Columns -------------------------------- */

  /**
   * Writes the structure into the file: deleted columns removed, new ones added, the order and names
   * rewritten, then the drop-down lists and colors, the look, and the app's data (types, views).
   */
  const sync = () =>
    queued("syncing", async () => {
      const e = entryOf();
      let t = await readTable(e.excel);
      const gone = (e.removed ?? []).filter((r) => t.header.includes(r));
      if (gone.length) await deleteColumns(e.excel, gone);
      const missing = e.columns.filter((c) => !fileName(c, t.header));
      if (missing.length) await addColumns(e.excel, missing.map((c) => c.name));
      t = await readTable(e.excel);
      // The app's order and names (the file's own extra columns, if any, kept after them).
      const order = [...e.columns.map((c) => fileName(c, t.header)!).filter(Boolean), ...t.header.filter((h) => !e.columns.some((c) => fileName(c, t.header) === h))];
      const names = order.map((h) => e.columns.find((c) => fileName(c, t.header) === h)?.name ?? h);
      if (JSON.stringify(names) !== JSON.stringify(t.header)) {
        const at = order.map((h) => t.header.indexOf(h));
        await rewriteTable(e.excel, names, t.rows.map((r) => at.map((i) => r[i] ?? "")));
      }
      const columns = e.columns.map(({ was: _, ...c }) => c);
      await writeChoicesWhenSaved(e.excel, columns, names);
      await formatTable(e.excel, names.map((n) => LOOKS[columns.find((c) => c.name === n)?.type ?? "text"]));
      const at = Date.now();
      await writeAppData(e.excel, { columns, views: e.views, at } satisfies AppData);
      await patch("tables", e.id, { columns, removed: [], changedAt: at });
      live.current.entry = { ...e, columns, removed: [], changedAt: at };
      await read();
    });

  /** Changes the columns here, then writes them to the file (a short pause groups quick changes). */
  const timer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const setColumns = async (columns: Column[], removed?: string[]) => {
    const e = entryOf();
    const next = { columns, removed: [...new Set([...(e.removed ?? []), ...(removed ?? [])])], changedAt: Date.now() };
    live.current.entry = { ...e, ...next };
    await patch("tables", e.id, next);
    setDirty(true);
    clearTimeout(timer.current);
    return new Promise<void>((resolve, reject) => (timer.current = setTimeout(() => void sync().then(resolve, reject), 1200)));
  };

  /** Views are kept here and in the file's app sheet (shared with the other users of the file). */
  const viewTimer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const setViews = async (views: View[]) => {
    const e = entryOf();
    const at = Date.now();
    live.current.entry = { ...e, views, changedAt: at };
    await patch("tables", e.id, { views, changedAt: at });
    clearTimeout(viewTimer.current);
    viewTimer.current = setTimeout(() => void writeAppData(e.excel, { columns: live.current.entry!.columns, views, at } satisfies AppData).catch(() => undefined), 2000);
  };

  return { data, status, dirty, reload, update, add, remove, sync, setColumns, setViews };
}
