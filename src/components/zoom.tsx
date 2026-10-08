import { useEffect, useState } from "react";
import { MaximizeIcon, MinusIcon, PlusIcon } from "lucide-react";
import { models, type Report } from "@/lib/pbi";

const STEPS = [0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4];

/** Zoom of an embedded report (Power BI hides its own zoom slider when embedded). */
export function Zoom({ report }: { report?: Report }) {
  const [zoom, setZoom] = useState<number>();
  useEffect(() => void report?.getZoom().then(setZoom, () => setZoom(undefined)), [report]);
  if (!report) return null;
  const go = async (next: number) => {
    await report.setZoom(next);
    setZoom(next);
  };
  const step = (dir: 1 | -1) => {
    const current = zoom ?? 1;
    const next = dir > 0 ? STEPS.find((s) => s > current + 0.01) : [...STEPS].reverse().find((s) => s < current - 0.01);
    if (next) void go(next);
  };
  const fit = async () => {
    const m = await models();
    await report.updateSettings({ layoutType: m.LayoutType.Custom, customLayout: { displayOption: m.DisplayOption.FitToPage } });
    await go(1);
  };
  return (
    <div className="flex items-center rounded-md border bg-card">
      <button type="button" className="icon-btn" title="Zoom out" onClick={() => step(-1)}>
        <MinusIcon />
      </button>
      <button type="button" className="h-7 w-12 text-[12px] tabular-nums hover:bg-muted" title="100 %" onClick={() => void go(1)}>
        {Math.round((zoom ?? 1) * 100)} %
      </button>
      <button type="button" className="icon-btn" title="Zoom in" onClick={() => step(1)}>
        <PlusIcon />
      </button>
      <button type="button" className="icon-btn" title="Fit to page" onClick={() => void fit()}>
        <MaximizeIcon />
      </button>
    </div>
  );
}
