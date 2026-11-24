/** A semantic model's structure (tables, columns with value examples, measures with DAX, relationships), downloadable as JSON. */
import { useEffect, useState } from "react";
import { DownloadIcon, RefreshCwIcon, TableIcon } from "lucide-react";
import type { Ref } from "@shared/types";
import { Empty, ErrorBox, Loading } from "@/components/common";
import { Button } from "@/components/ui/button";
import { download, errorText, get } from "@/lib/api";

type Info = { source: string; tables: { name: string; description?: string; hidden?: boolean; kind?: string; rows?: number; columns: { name: string; type?: string; values?: string; hidden?: boolean; description?: string; expression?: string }[]; measures: { name: string; expression?: string; format?: string; folder?: string }[]; hierarchies: string[] }[]; relationships: { from: string; to: string; cardinality: string; both: boolean; active: boolean }[]; roles: string[]; warnings: string[]; readAt: string };

export function useModel(ref?: Ref, reload = 0) {
  const [info, setInfo] = useState<Info>();
  const [error, setError] = useState<string>();
  useEffect(() => {
    if (!ref?.id) return;
    setInfo(undefined);
    setError(undefined);
    const q = new URLSearchParams({ id: ref.id, ...(ref.groupId ? { groupId: ref.groupId } : {}), ...(reload ? { fresh: "1" } : {}) });
    get<Info>(`/powerbi/model?${q}`).then(setInfo, (e) => setError(errorText(e)));
  }, [ref?.id, ref?.groupId, reload]);
  return { info, error };
}

export function ModelView({ refOf, name }: { refOf?: Ref; name?: string }) {
  const [reload, setReload] = useState(0);
  const { info, error } = useModel(refOf, reload);
  const [open, setOpen] = useState<string>();
  if (!refOf) return <Empty icon={<TableIcon />} title="No semantic model" />;
  if (error) return <div className="p-6"><ErrorBox>{error}</ErrorBox></div>;
  if (!info) return <Loading label="Reading the model (structure and value examples)…" />;
  const tables = info.tables.filter((t) => !t.hidden || t.measures.length);
  return (
    <div className="flex flex-col gap-4 p-6">
      <div className="flex flex-wrap items-center gap-2">
        <div className="flex-1 text-[13px] text-muted-foreground">
          {info.tables.length} tables · {info.tables.reduce((n, t) => n + t.measures.length, 0)} measures · {info.relationships.length} relationships · {info.source === "tmdl" ? "full definition" : "read with DAX"}
        </div>
        <Button variant="outline" size="sm" onClick={() => setReload((n) => n + 1)}>
          <RefreshCwIcon /> Read again
        </Button>
        <Button variant="dark" size="sm" onClick={() => download(new Blob([JSON.stringify(info, null, 2)], { type: "application/json" }), `${name ?? "model"}.json`)}>
          <DownloadIcon /> Structure (JSON)
        </Button>
      </div>
      {info.warnings.map((w) => (
        <div key={w} className="rounded-xl bg-warning/10 px-3 py-2 text-[12.5px] text-warning">{w}</div>
      ))}
      <div className="grid gap-3 xl:grid-cols-2">
        {tables.map((t) => (
          <div key={t.name} className="rounded-2xl border">
            <button onClick={() => setOpen(open === t.name ? undefined : t.name)} className="flex w-full items-center gap-2 px-4 py-3 text-left">
              <TableIcon className="size-4 text-muted-foreground" />
              <span className="flex-1 font-semibold">{t.name}</span>
              <span className="text-[12px] text-muted-foreground">
                {t.rows !== undefined ? `${t.rows.toLocaleString()} rows · ` : ""}
                {t.columns.length} col. · {t.measures.length} meas.
              </span>
            </button>
            {open === t.name && (
              <div className="flex flex-col gap-1 border-t px-4 py-3 text-[12.5px]">
                {t.description && <div className="mb-1 text-muted-foreground">{t.description}</div>}
                {t.columns.filter((c) => !c.hidden).map((c) => (
                  <div key={c.name} className="flex gap-2">
                    <span className="font-medium">{c.name}</span>
                    <span className="text-muted-foreground">{c.type}</span>
                    <span className="min-w-0 flex-1 truncate text-right text-muted-foreground" title={c.values}>{c.values}</span>
                  </div>
                ))}
                {t.measures.map((m) => (
                  <details key={m.name} className="rounded-lg bg-muted/50 px-2 py-1">
                    <summary className="font-medium text-primary">[{m.name}]</summary>
                    <pre className="mt-1 whitespace-pre-wrap font-mono text-[11.5px]">{m.expression ?? "(formula not readable)"}</pre>
                  </details>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
      {info.relationships.length > 0 && (
        <div className="rounded-2xl border p-4 text-[12.5px]">
          <div className="mb-2 font-semibold">Relationships</div>
          {info.relationships.map((r, i) => (
            <div key={i} className="text-muted-foreground">
              {r.from} <span className="text-foreground">{r.cardinality}</span> {r.to}
              {r.both ? " · both directions" : ""}
              {r.active ? "" : " · inactive"}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
