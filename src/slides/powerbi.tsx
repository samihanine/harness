/**
 * A Power BI report page (or a zone of it) as a plain iframe ("reportEmbed" with the browser's Power BI
 * sign-in): no SDK, so it also prints into the PDF export. The page is drawn at 1280×720 and the zone
 * (`crop`, in page px) is scaled to fill its box.
 */
import { useEffect, useRef, useState } from "react";
import { z } from "zod";
import { useCollection } from "@/lib/store";

export const PAGE = { w: 1280, h: 720 };
export const cropSchema = z.object({ x: z.number(), y: z.number(), w: z.number(), h: z.number() }).describe("Zone of the page in px of a 1280×720 page");
export type Crop = z.infer<typeof cropSchema>;

/** The report of the app given by its id or its name. */
export function useReportRef(report: string) {
  const reports = useCollection("reports");
  return reports.find((r) => r.id === report) ?? reports.find((r) => r.name.toLowerCase() === report.toLowerCase());
}

export const embedUrl = (r: { id: string; groupId?: string }, page?: string) =>
  `https://app.powerbi.com/reportEmbed?reportId=${r.id}${r.groupId ? `&groupId=${r.groupId}` : ""}&autoAuth=true&navContentPaneEnabled=false&filterPaneEnabled=false${page ? `&pageName=${encodeURIComponent(page)}` : ""}`;

export function PowerBiFrame({ report, page, crop }: { report: string; page?: string; crop?: Crop }) {
  const ref = useReportRef(report);
  const box = useRef<HTMLDivElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });
  useEffect(() => {
    const el = box.current;
    if (!el) return;
    const observer = new ResizeObserver(([e]) => setSize({ w: e.contentRect.width, h: e.contentRect.height }));
    observer.observe(el);
    return () => observer.disconnect();
  }, []);
  const zone = crop ?? { x: 0, y: 0, ...PAGE };
  const scale = size.w && size.h ? Math.min(size.w / zone.w, size.h / zone.h) : 0;
  return (
    <div ref={box} className="relative size-full overflow-hidden rounded-[0.8cqw]" style={{ background: "color-mix(in srgb, var(--s-text) 4%, var(--s-bg))" }}>
      {!ref ? (
        <div className="flex size-full items-center justify-center p-[2cqw] text-center text-[1.2cqw] opacity-60">Power BI report “{report}” is not in the app (add it in Reports).</div>
      ) : (
        scale > 0 && (
          <div className="absolute top-1/2 left-1/2" style={{ width: zone.w * scale, height: zone.h * scale, transform: "translate(-50%, -50%)", overflow: "hidden" }}>
            <iframe
              title={`Power BI ${ref.name}`}
              src={embedUrl(ref, page)}
              className="absolute border-0"
              style={{ width: PAGE.w, height: PAGE.h + 36, left: 0, top: 0, transformOrigin: "0 0", transform: `scale(${scale}) translate(${-zone.x}px, ${-zone.y}px)` }}
            />
          </div>
        )
      )}
    </div>
  );
}
