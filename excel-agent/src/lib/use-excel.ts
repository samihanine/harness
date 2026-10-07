import { useCallback, useEffect, useState, useSyncExternalStore } from "react";
import { PermissionNeededError, SignInRequiredError, openStorage } from "@/storage/open";
import type { ImageStore } from "@/storage/types";
import { ExcelFile } from "./excel";
import type { Schema } from "./schema";
import type { FileLink } from "./store";
import { links } from "./store";

type State =
  | { status: "loading" }
  | { status: "permission"; fileName: string }
  | { status: "signin" }
  | { status: "error"; error: string }
  | { status: "ready"; excel: ExcelFile; images?: ImageStore };

/** Opens the file of a link through its storage and follows its changes. */
export function useExcel(link: FileLink | undefined, schema: Schema | undefined) {
  const [state, setState] = useState<State>({ status: "loading" });

  const open = useCallback(
    async (ask: boolean) => {
      if (!link || !schema) return setState({ status: "error", error: "File or schema not found." });
      // The stored link, not the one of this render: a reopen right after a change (images
      // folder…) would otherwise open, and save back, the previous version of the link.
      const current = links.get(link.id) ?? link;
      try {
        const { backend, images } = await openStorage(current, ask);
        const excel = await ExcelFile.open(backend, schema);
        links.put({ ...(links.get(link.id) ?? current), openedAt: Date.now() });
        setState({ status: "ready", excel, images });
      } catch (error) {
        if (error instanceof PermissionNeededError) setState({ status: "permission", fileName: error.fileName });
        else if (error instanceof SignInRequiredError) setState({ status: "signin" });
        else setState({ status: "error", error: error instanceof Error ? error.message : String(error) });
      }
    },
    // Reopen when the schema or the storage changes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [link?.id, JSON.stringify(schema), JSON.stringify(link?.source), JSON.stringify(link?.images)],
  );

  useEffect(() => {
    void open(false);
  }, [open]);

  const excel = state.status === "ready" ? state.excel : undefined;
  useSyncExternalStore(
    (listener) => excel?.subscribe(listener) ?? (() => {}),
    () => excel?.version ?? -1,
  );
  // Follow changes made outside (Excel, Excel online, colleagues…).
  useEffect(() => {
    if (!excel) return;
    const timer = setInterval(() => void excel.reloadIfChanged().catch(() => undefined), excel.backend.pollEvery);
    return () => clearInterval(timer);
  }, [excel]);

  return { state, reopen: () => open(true) };
}
