/** A report's structure: pages and their visuals with fields and filters; downloadable as JSON (all formatting when editable). */
import { useEffect, useState } from "react";
import { DownloadIcon, FileJsonIcon } from "lucide-react";
import type { ReportRef } from "@shared/types";
import { Empty, ErrorBox, Loading } from "@/components/common";
import { Button } from "@/components/ui/button";
import { download, errorText, get } from "@/lib/api";
import { toast } from "sonner";

type Structure = { source: string; note?: string; pages: { name: string; displayName: string; hidden?: boolean; filters?: unknown[]; visuals: { name: string; type: string; title?: string; fields?: Record<string, string[]>; slicerOn?: string[]; filters?: unknown[]; position?: { x: number; y: number; width: number; height: number } }[] }[]; reportFilters?: unknown[]; theme?: { base?: string; custom?: string }; measures?: { table: string; name: string; expression: string }[] };

const query = (r: ReportRef, detail = "summary") => new URLSearchParams({ id: r.id, detail, ...(r.groupId ? { groupId: r.groupId } : {}), ...(r.appId ? { appId: r.appId } : {}) });

export function StructureView({ report }: { report?: ReportRef }) {
  const [s, setS] = useState<Structure>();
  const [error, setError] = useState<string>();
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    if (!report) return;
    setS(undefined);
    setError(undefined);
    get<Structure>(`/powerbi/report-structure?${query(report)}`).then(setS, (e) => setError(errorText(e)));
  }, [report?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!report) return <Empty icon={<FileJsonIcon />} title="No report" />;
  if (error) return <div className="p-6"><ErrorBox>{error}</ErrorBox></div>;
  if (!s) return <Loading label="Reading the report's structure…" />;
  const dl = async () => {
    setBusy(true);
    try {
      const full = await get(`/powerbi/report-structure?${query(report, "full")}`);
      download(new Blob([JSON.stringify(full, null, 2)], { type: "application/json" }), `${report.name}.json`);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  return (
    <div className="flex flex-col gap-4 p-6">
      <div className="flex flex-wrap items-center gap-2">
        <div className="flex-1 text-[13px] text-muted-foreground">
          {s.pages.length} pages · {s.pages.reduce((n, p) => n + p.visuals.length, 0)} visuals · {s.source === "definition" ? `editable — theme ${s.theme?.custom ?? s.theme?.base ?? "default"}, ${s.measures?.length ?? 0} report measures` : "view only (read through an embed)"}
        </div>
        <Button variant="dark" size="sm" onClick={dl} disabled={busy}>
          <DownloadIcon /> Structure (JSON)
        </Button>
      </div>
      {s.pages.map((p) => (
        <div key={p.name} className="rounded-2xl border">
          <div className="flex items-center gap-2 border-b px-4 py-3">
            <span className="flex-1 font-semibold">{p.displayName}</span>
            {p.hidden && <span className="badge bg-muted">hidden</span>}
            <span className="text-[12px] text-muted-foreground">{p.visuals.length} visuals</span>
          </div>
          <div className="divide-y">
            {p.visuals.map((v) => (
              <div key={v.name} className="flex gap-3 px-4 py-2 text-[12.5px]">
                <span className="w-36 shrink-0 truncate font-mono text-muted-foreground">{v.type}</span>
                <span className="w-48 shrink-0 truncate font-medium">{v.title ?? ""}</span>
                <span className="min-w-0 flex-1 truncate text-muted-foreground">
                  {v.fields ? Object.entries(v.fields).map(([r, f]) => `${r}: ${f.join(", ")}`).join(" · ") : v.slicerOn?.join(", ")}
                </span>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
