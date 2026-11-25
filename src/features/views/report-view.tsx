/** An embedded Power BI report filling its area, with zoom and "fit" (it fits by itself when it opens). */
import { useEffect, useState } from "react";
import { Loader2Icon, MinusIcon, PlusIcon, ScanIcon } from "lucide-react";
import type { Ref } from "@shared/types";
import { ErrorBox } from "@/components/common";
import { Button } from "@/components/ui/button";
import { fitReport, type Report, type useEmbed } from "@/features/powerbi/embed";

const STEPS = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3];

function Zoom({ report, element }: { report: Report; element: HTMLElement | null }) {
  const [zoom, setZoom] = useState<number>(1);
  useEffect(() => void report.getZoom().then(setZoom, () => undefined), [report]);
  const go = async (next: number) => {
    await report.setZoom(next).catch(() => undefined);
    setZoom(next);
  };
  const step = (dir: 1 | -1) => {
    const next = dir > 0 ? STEPS.find((s) => s > zoom + 0.01) : [...STEPS].reverse().find((s) => s < zoom - 0.01);
    if (next) void go(next);
  };
  return (
    <div className="absolute right-6 bottom-6 z-10 flex h-9 items-center gap-0.5 rounded-full border bg-card/95 px-1 shadow-sm backdrop-blur">
      <Button variant="ghost" size="icon-sm" aria-label="Zoom out" onClick={() => step(-1)}><MinusIcon /></Button>
      <button className="h-7 w-12 rounded-full text-[12px] tabular-nums hover:bg-muted" onClick={() => void go(1)} aria-label="Zoom 100 %">{Math.round(zoom * 100)} %</button>
      <Button variant="ghost" size="icon-sm" aria-label="Zoom in" onClick={() => step(1)}><PlusIcon /></Button>
      <Button variant="ghost" size="icon-sm" aria-label="Fit to the screen" title="Fit to the screen" onClick={() => void fitReport(report, element).then(() => setZoom(1))}><ScanIcon /></Button>
    </div>
  );
}

export function ReportFrame({ embed, refOf }: { embed: ReturnType<typeof useEmbed>; refOf?: Ref }) {
  // Fitted again when the space changes (assistant folded, window resized).
  useEffect(() => {
    const el = embed.element.current;
    if (!el || !embed.report) return;
    let t: ReturnType<typeof setTimeout>;
    const ro = new ResizeObserver(() => (clearTimeout(t), (t = setTimeout(() => void fitReport(embed.report!, el), 300))));
    ro.observe(el);
    return () => (ro.disconnect(), clearTimeout(t));
  }, [embed.report]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div className="relative min-h-0 flex-1 p-3">
      <div ref={embed.element} className="h-full w-full overflow-hidden rounded-2xl bg-muted/40 [&_iframe]:border-0" />
      {embed.report && <Zoom report={embed.report} element={embed.element.current} />}
      {refOf && embed.loading && (
        <div className="absolute inset-0 flex items-center justify-center gap-2 text-[13px] text-muted-foreground">
          <Loader2Icon className="size-4 animate-spin" /> Loading the report…
        </div>
      )}
      {embed.error && (
        <div className="absolute inset-x-6 top-6">
          <ErrorBox>{embed.error}</ErrorBox>
        </div>
      )}
    </div>
  );
}
