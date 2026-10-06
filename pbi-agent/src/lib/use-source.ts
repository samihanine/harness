import { useEffect, useState } from "react";
import { useDb } from "@/db/db";
import type { Source } from "@/query/engine";
import { loadSource } from "@/query/engine";

/** Loads the sources (datasets with their model) for the given dataset ids. */
export function useSources(ids: string[]) {
  const data = useDb();
  const [sources, setSources] = useState<Source[]>([]);
  const [error, setError] = useState("");
  const key = ids.join(",") + ":" + ids.map((id) => data.row("datasets", id)?.content_json).join(",");
  useEffect(() => {
    let cancelled = false;
    Promise.all(ids.flatMap((id) => (data.row("datasets", id) ? [loadSource(data.row("datasets", id)!)] : [])))
      .then((all) => !cancelled && (setSources(all), setError("")))
      .catch((e: Error) => !cancelled && setError(e.message));
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key]);
  return { sources, error };
}
