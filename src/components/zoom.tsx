import { useEffect, useState } from "react";
import { MinusIcon, MoveHorizontalIcon, PlusIcon } from "lucide-react";
import { Button } from "./ui/button";
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
  /** Page as wide as the frame (scrolls vertically), not the whole page centered. */
  const fit = async () => {
    const m = await models();
    await report.setZoom(1).catch(() => undefined);
    setZoom(1);
    await report.updateSettings({ layoutType: m.LayoutType.Custom, customLayout: { displayOption: m.DisplayOption.FitToWidth } });
  };
  return (
    <div className="flex h-8 items-center rounded-md border border-input bg-card shadow-xs">
      <Button variant="ghost" size="icon" className="size-7" aria-label="Zoom out" onClick={() => step(-1)}>
        <MinusIcon />
      </Button>
      <button type="button" className="h-7 w-12 rounded text-[12px] tabular-nums hover:bg-accent" aria-label="Zoom 100 %" onClick={() => void go(1)}>
        {Math.round((zoom || 1) * 100)} %
      </button>
      <Button variant="ghost" size="icon" className="size-7" aria-label="Zoom in" onClick={() => step(1)}>
        <PlusIcon />
      </Button>
      <Button variant="ghost" size="icon" className="size-7" aria-label="Fit to width" onClick={() => void fit()}>
        <MoveHorizontalIcon />
      </Button>
    </div>
  );
}
