import { useCallback, useEffect, useState, useSyncExternalStore } from "react";
import { ExcelFile } from "./excel";
import type { Schema } from "./schema";
import type { FileLink } from "./store";
import { handles, links, permission } from "./store";

type State =
  | { status: "loading" }
  | { status: "permission"; fileName: string }
  | { status: "error"; error: string }
  | { status: "ready"; excel: ExcelFile; images?: FileSystemDirectoryHandle };

const RELOAD_EVERY = 3_000;

/** Opens the local file of a link (asking for access when needed) and follows its changes. */
export function useExcel(link: FileLink | undefined, schema: Schema | undefined) {
  const [state, setState] = useState<State>({ status: "loading" });

  const open = useCallback(
    async (ask: boolean) => {
      if (!link || !schema) return setState({ status: "error", error: "File or schema not found." });
      try {
        const file = await handles.file(link.id);
        if (!file) return setState({ status: "error", error: "The file handle is missing: add the file again." });
        const images = await handles.images(link.id);
        if (!(await permission(file, ask)) || (images && !(await permission(images, ask))))
          return setState({ status: "permission", fileName: link.fileName });
        const excel = await ExcelFile.open(file, schema);
        links.put({ ...link, openedAt: Date.now() });
        setState({ status: "ready", excel, images });
      } catch (error) {
        setState({ status: "error", error: error instanceof Error ? error.message : String(error) });
      }
    },
    // Reopen when the schema changes (new columns…).
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [link?.id, JSON.stringify(schema)],
  );

  useEffect(() => {
    void open(false);
  }, [open]);

  const excel = state.status === "ready" ? state.excel : undefined;
  useSyncExternalStore(
    (listener) => excel?.subscribe(listener) ?? (() => {}),
    () => excel?.version ?? -1,
  );
  // Follow changes made outside (Excel…).
  useEffect(() => {
    if (!excel) return;
    const timer = setInterval(() => void excel.reloadIfChanged(), RELOAD_EVERY);
    return () => clearInterval(timer);
  }, [excel]);

  return { state, grant: () => open(true) };
}
