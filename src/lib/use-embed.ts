import { useEffect, useRef, useState } from "react";
import { errorText } from "./ms";
import { embedReport, resetEmbed, type EmbedMode, type Ref, type Report } from "./pbi";

/** Embeds a report in the returned element; follows the current page and the clicked visual. */
export function useEmbed(ref: Ref | undefined, mode: EmbedMode) {
  const element = useRef<HTMLDivElement>(null);
  const [report, setReport] = useState<Report>();
  const [error, setError] = useState<string>();
  const [selection, setSelection] = useState<{ page?: string; visual?: string }>({});
  useEffect(() => {
    const el = element.current;
    if (!ref || !el) return;
    let live = true;
    setReport(undefined);
    setError(undefined);
    embedReport(el, ref, mode).then(
      async (r) => {
        if (!live) return;
        r.on("pageChanged", (e: any) => setSelection({ page: e.detail.newPage.name }));
        r.on("visualClicked", (e: any) => setSelection({ page: e.detail.page?.name, visual: e.detail.visual?.name }));
        const page = (await r.getPages()).find((p) => p.isActive);
        setSelection({ page: page?.name });
        setReport(r);
      },
      (e) => live && setError(errorText(e)),
    );
    return () => {
      live = false;
      void resetEmbed(el);
    };
  }, [ref?.id, ref?.groupId, mode]); // eslint-disable-line react-hooks/exhaustive-deps
  return { element, report, error, selection };
}
