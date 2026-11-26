/** An embedded Power BI report filling its area, with zoom and "fit to the width" (its 100 %, set when it opens and at each page change). */
import { useEffect, useState } from "react";
import { Loader2Icon, MinusIcon, MoveHorizontalIcon, PlusIcon } from "lucide-react";
import type { Ref } from "@shared/types";
import { ErrorBox } from "@/components/common";
import { Button } from "@/components/ui/button";
import { fitReport, type Report, type useEmbed } from "@/features/powerbi/embed";

/** Zoom steps relative to the fitted width. */
const STEPS = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3];

function Zoom({ report, element }: { report: Report; element: HTMLElement | null }) {
  // `base`: Power BI's zoom when the page fits the width (shown as 100 %); `zoom`: the zoom now.
  const [base, setBase] = useState(1);
  const [zoom, setZoom] = useState(1);
  const fit = () => fitReport(report).then((z) => (setBase(z), setZoom(z)));
  useEffect(() => {
    // Each page opens fitted to the width (also after a resize of the area).
    const refit = () => void fit();
    refit();
    report.on("pageChanged", refit);
    let t: ReturnType<typeof setTimeout>;
    const ro = new ResizeObserver(() => (clearTimeout(t), (t = setTimeout(refit, 300))));
    if (element) ro.observe(element);
    return () => (report.off("pageChanged", refit), ro.disconnect(), clearTimeout(t));
  }, [report, element]); // eslint-disable-line react-hooks/exhaustive-deps
  const relative = zoom / base;
  const step = (dir: 1 | -1) => {
    const next = dir > 0 ? STEPS.find((s) => s > relative + 0.01) : [...STEPS].reverse().find((s) => s < relative - 0.01);
    if (!next) return;
    if (Math.abs(next - 1) < 0.01) return void fit();
    void report.setZoom(next * base).then(() => setZoom(next * base), () => undefined);
  };
  return (
    <div className="absolute right-6 bottom-6 z-10 flex h-9 items-center gap-0.5 rounded-full border bg-card/95 px-1 shadow-sm backdrop-blur">
      <Button variant="ghost" size="icon-sm" aria-label="Zoom out" onClick={() => step(-1)}><MinusIcon /></Button>
      <button className="h-7 w-12 rounded-full text-[12px] tabular-nums hover:bg-muted" onClick={() => void fit()} aria-label="Zoom 100 % (the report's width)">{Math.round(relative * 100)} %</button>
      <Button variant="ghost" size="icon-sm" aria-label="Zoom in" onClick={() => step(1)}><PlusIcon /></Button>
      <Button variant="ghost" size="icon-sm" aria-label="Fit to the width" title="Fit to the width" onClick={() => void fit()}><MoveHorizontalIcon /></Button>
    </div>
  );
}

export function ReportFrame({ embed, refOf }: { embed: ReturnType<typeof useEmbed>; refOf?: Ref }) {
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
