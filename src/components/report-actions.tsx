/** Open the report in Power BI, and export it (PDF through the print dialog, or .pbix). */
import { useState } from "react";
import { DownloadIcon, ExternalLinkIcon } from "lucide-react";
import { errorText } from "@/lib/ms";
import type { Report } from "@/lib/pbi";
import type { ReportEntry } from "@/lib/store";
import { exportPbix } from "@/server/ms";
import { Popover } from "./popover";

/**
 * The Power BI "Export to file" API (PDF, PowerPoint, PNG) needs a Premium / Fabric capacity: on Pro,
 * PDF goes through the browser print dialog ("Save as PDF"), and the .pbix download works.
 */
export function ReportActions({ entry, report }: { entry?: ReportEntry; report?: Report }) {
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string>();
  if (!entry) return null;
  const url = entry.appId ? `https://app.powerbi.com/groups/me/apps/${entry.appId}/reports/${entry.id}` : `https://app.powerbi.com/groups/${entry.groupId ?? "me"}/reports/${entry.id}`;
  const pbix = async (close: () => void) => {
    setBusy(true);
    setError(undefined);
    try {
      const base64 = await exportPbix({ data: { id: entry.id, groupId: entry.groupId } });
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(new Blob([bytes])), download: `${entry.name}.pbix` });
      a.click();
      URL.revokeObjectURL(a.href);
      close();
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  return (
    <>
      <a href={url} target="_blank" rel="noreferrer" className="btn h-8" title="Open this report on the Power BI site">
        <ExternalLinkIcon className="size-3.5" /> Power BI
      </a>
      <Popover
        width={260}
        trigger={(open) => (
          <button type="button" className="btn h-8" onClick={() => (setError(undefined), open())}>
            <DownloadIcon className="size-3.5" /> Export
          </button>
        )}
      >
        {(close) => (
          <div className="flex flex-col">
            <button type="button" disabled={!report} onClick={() => (close(), void report?.print())} className="rounded-md px-2 py-1.5 text-left hover:bg-muted disabled:opacity-50">
              PDF of the current page
              <span className="block text-[11px] text-muted-foreground">Print dialog → Save as PDF</span>
            </button>
            {!entry.appId && <button type="button" disabled={busy} onClick={() => void pbix(close)} className="rounded-md px-2 py-1.5 text-left hover:bg-muted disabled:opacity-50">
              {busy ? "Downloading…" : "Power BI file (.pbix)"}
              <span className="block text-[11px] text-muted-foreground">Opens in Power BI Desktop, connected to the online model</span>
            </button>}
            {error && <p className="px-2 pt-1 text-[11px] text-destructive">{error}</p>}
          </div>
        )}
      </Popover>
    </>
  );
}
