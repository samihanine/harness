import { useCallback, useEffect, useRef, useState } from "react";
import { readRows, type Table } from "./excel";
import { errorText } from "./ms";
import type { SchemaEntry } from "./store";

/**
 * The rows of a schema's Excel table. Writes are optimistic: the change is shown at once, the requests
 * run one after the other (row positions stay consistent), and a failure restores the file's content.
 */
export function useTable(schema?: SchemaEntry) {
  const [table, setTable] = useState<Table>();
  const [error, setError] = useState<string>();
  const current = useRef<Table | undefined>(undefined);
  current.current = table;
  const queue = useRef<Promise<unknown>>(Promise.resolve());
  const pending = useRef({ count: 0, refetch: false });
  const show = (t?: Table) => {
    current.current = t;
    setTable(t);
  };
  const excel = schema?.excel;
  const key = excel ? `${excel.itemId}|${excel.table}|${JSON.stringify(schema!.fields)}` : undefined;
  const id = useRef(key);
  id.current = key;
  const reload = useCallback(async () => {
    if (!excel || !schema) return show(undefined);
    try {
      const t = await readRows(excel, schema.fields);
      // Another file was chosen meanwhile: its own read shows it.
      if (id.current !== key) return;
      show(t);
      setError(undefined);
    } catch (e) {
      if (id.current === key) setError(errorText(e));
    }
  }, [key]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    // Not the previous file's rows (nor its error) while the new one loads.
    show(undefined);
    setError(undefined);
    void reload();
  }, [reload]);

  const write = useCallback(
    (task: (t: Table) => Promise<unknown>, { refetch = false, optimistic }: { refetch?: boolean; optimistic?: (t: Table) => Table } = {}) => {
      const base = current.current;
      if (!base) return Promise.resolve();
      let shown = base;
      try {
        if (optimistic) show((shown = optimistic(base)));
      } catch (e) {
        setError(errorText(e)); // invalid value: nothing sent
        return Promise.reject(e);
      }
      pending.current.count++;
      pending.current.refetch ||= refetch;
      const run = queue.current.then(async () => {
        try {
          await task(optimistic ? { ...shown, rows: [...shown.rows] } : shown);
          setError(undefined);
          if (!optimistic) show({ ...shown, rows: [...shown.rows] });
        } catch (e) {
          setError(errorText(e));
          pending.current.refetch = true; // the real content is read back
          throw e;
        } finally {
          // Read back once the queue is empty (never over a pending optimistic change).
          if (--pending.current.count === 0 && pending.current.refetch) {
            pending.current.refetch = false;
            await reload();
          }
        }
      });
      queue.current = run.catch(() => undefined);
      return run;
    },
    [reload],
  );
  return { table, error, reload, write };
}

export type Write = ReturnType<typeof useTable>["write"];
