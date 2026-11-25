/**
 * A report's structure: pages and their visuals with fields and filters — kept in this browser (read when
 * the report was added, "Refresh" reads it again), downloadable as JSON (all formatting when editable).
 */
import { useEffect, useState } from "react";
import { DownloadIcon, FileJsonIcon, Loader2Icon, RefreshCwIcon } from "lucide-react";
import { toast } from "sonner";
import type { ReportRef } from "@shared/types";
import { Empty, ErrorBox, Loading } from "@/components/common";
import { Button } from "@/components/ui/button";
import { download, errorText, get } from "@/lib/api";
import { reportStructure } from "@/lib/items";
import { useItems } from "@/lib/store";

type Filter = { field: string; type?: string; selection?: string; hidden?: boolean };
type Visual = { name: string; type: string; title?: string; fields?: Record<string, string[]>; slicerOn?: string[]; filters?: (Filter | string)[]; hidden?: boolean };
type Structure = { source: string; note?: string; format?: string; reportFilters?: Filter[]; pages: { name: string; displayName: string; hidden?: boolean; filters?: Filter[]; visuals: Visual[] }[]; theme?: { base?: string; custom?: string }; measures?: unknown[] };

/** Filters as chips: the field, and what it keeps. */
function Filters({ label, list }: { label: string; list?: (Filter | string)[] }) {
  if (!list?.length) return null;
  return (
    <div className="flex flex-wrap items-baseline gap-1.5 text-[12px]">
      <span className="text-muted-foreground">{label}</span>
      {list.map((f, i) => {
        const x = typeof f === "string" ? { field: f } : f;
        return (
          <span key={i} className="rounded-md bg-muted px-1.5 py-0.5 font-mono text-[11.5px]">
            {x.field}
            {x.selection ? <span className="text-muted-foreground"> {x.selection}</span> : null}
            {x.hidden ? <span className="text-muted-foreground"> (hidden)</span> : null}
          </span>
        );
      })}
    </div>
  );
}

const query = (r: ReportRef, detail: string) => new URLSearchParams({ id: r.id, detail, ...(r.groupId ? { groupId: r.groupId } : {}), ...(r.appId ? { appId: r.appId } : {}) });

export function StructureView({ report }: { report?: ReportRef }) {
  const kept = (useItems("report") ?? []).find((r) => r.id === report?.id);
  const s = kept?.structure as Structure | undefined;
  const [error, setError] = useState<string>();
  const [busy, setBusy] = useState<string>();
  const read = (fresh: boolean) => {
    if (!report) return;
    setError(undefined);
    setBusy(fresh ? "refresh" : "read");
    reportStructure(report, fresh).then(() => setBusy(undefined), (e) => (setError(errorText(e)), setBusy(undefined)));
  };
  useEffect(() => {
    if (report && !kept?.structure) read(false);
  }, [report?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!report) return <Empty icon={<FileJsonIcon />} title="No report" />;
  if (error && !s) return <div className="p-6"><ErrorBox>{error}</ErrorBox></div>;
  if (!s) return <Loading label="Reading the report's structure…" />;
  const dl = async () => {
    setBusy("download");
    try {
      const full = s.source === "definition" ? await get(`/powerbi/report-structure?${query(report, "full")}`) : s;
      download(new Blob([JSON.stringify(full, null, 2)], { type: "application/json" }), `${report.name}.json`);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  return (
    <div className="flex flex-col gap-4 p-6">
      <div className="flex flex-wrap items-center gap-2">
        <div className="min-w-[240px] flex-1 text-[13px] text-muted-foreground">
          {s.pages.length} pages · {s.pages.reduce((n, p) => n + p.visuals.length, 0)} visuals · {s.source === "definition" ? `full definition${s.format ? ` (${s.format})` : ""} — theme ${s.theme?.custom ?? s.theme?.base ?? "default"}, ${s.measures?.length ?? 0} report measures` : "read through an embed (view only)"}
          {kept?.structureAt && ` · read ${new Date(kept.structureAt).toLocaleString()}`}
        </div>
        <Button variant="outline" size="sm" onClick={() => read(true)} disabled={!!busy}>
          {busy === "refresh" ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Refresh
        </Button>
        <Button variant="dark" size="sm" onClick={dl} disabled={!!busy}>
          {busy === "download" ? <Loader2Icon className="animate-spin" /> : <DownloadIcon />} Structure (JSON)
        </Button>
      </div>
      {error && <ErrorBox>{error}</ErrorBox>}
      {!!s.reportFilters?.length && (
        <div className="rounded-2xl border px-4 py-3">
          <Filters label="Report filters" list={s.reportFilters} />
        </div>
      )}
      {s.pages.map((p) => (
        <div key={p.name} className="rounded-2xl border">
          <div className="flex items-center gap-2 border-b px-4 py-3">
            <span className="flex-1 font-semibold">{p.displayName}</span>
            {p.hidden && <span className="badge bg-muted">hidden</span>}
            <span className="text-[12px] text-muted-foreground">{p.visuals.length} visuals</span>
          </div>
          {!!p.filters?.length && (
            <div className="border-b px-4 py-2">
              <Filters label="Page filters" list={p.filters} />
            </div>
          )}
          <div className="divide-y">
            {p.visuals.map((v) => (
              <div key={v.name} className="flex flex-wrap gap-x-3 gap-y-1 px-4 py-2 text-[12.5px]">
                <span className="w-36 min-w-0 truncate font-mono text-muted-foreground" title={v.name}>{v.type}</span>
                <span className="w-48 min-w-0 truncate font-medium" title={v.title}>{v.title ?? ""}{v.hidden ? <span className="font-normal text-muted-foreground"> (hidden)</span> : null}</span>
                <div className="flex min-w-[220px] flex-1 basis-[260px] flex-col gap-1">
                  {v.fields && Object.keys(v.fields).length ? (
                    Object.entries(v.fields).map(([r, f]) => (
                      <div key={r} className="break-words">
                        <span className="text-muted-foreground">{r}: </span>
                        <span className="font-mono text-[11.5px]">{f.join(", ")}</span>
                      </div>
                    ))
                  ) : v.slicerOn ? (
                    <div className="break-words"><span className="text-muted-foreground">Slicer on: </span><span className="font-mono text-[11.5px]">{v.slicerOn.join(" > ")}</span></div>
                  ) : (
                    <span className="text-muted-foreground">—</span>
                  )}
                  <Filters label="Filters" list={v.filters} />
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
