/** Open the report in Power BI, and export it (PDF through the print dialog, or .pbix). */
import { DownloadIcon, ExternalLinkIcon, FileDownIcon, PrinterIcon } from "lucide-react";
import { Button } from "./ui/button";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "./ui/dropdown-menu";
import { withToast } from "./entity";
import { EXPORT } from "@/config";
import type { Report } from "@/services/pbi";
import type { ReportEntry } from "@/lib/store";
import { exportPbix } from "@/services/ms";

/** Downloads a report as a .pbix file (it opens in Power BI Desktop, connected to the online model). */
export async function downloadPbix(report: { id: string; groupId?: string; name: string }) {
  const blob = await exportPbix({ id: report.id, groupId: report.groupId });
  const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: `${report.name}.pbix` });
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10_000);
}

/**
 * The Power BI "Export to file" API (PDF, PowerPoint, PNG) needs a Premium / Fabric capacity: on Pro,
 * PDF goes through the browser print dialog ("Save as PDF"), and the .pbix download works.
 */
export function ReportActions({ entry, report }: { entry?: ReportEntry; report?: Report }) {
  if (!entry) return null;
  const url = entry.appId ? `https://app.powerbi.com/groups/me/apps/${entry.appId}/reports/${entry.id}` : `https://app.powerbi.com/groups/${entry.groupId ?? "me"}/reports/${entry.id}`;
  return (
    <>
      <Button variant="outline" size="sm" asChild>
        <a href={url} target="_blank" rel="noreferrer">
          <ExternalLinkIcon /> Power BI
        </a>
      </Button>
      {EXPORT.report && (
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button variant="outline" size="sm">
              <DownloadIcon /> Export
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="w-64">
            <DropdownMenuItem disabled={!report} onSelect={() => void report?.print()}>
              <PrinterIcon />
              <span className="flex flex-col">
                PDF of the current page
                <span className="text-[11px] text-muted-foreground">Print dialog → Save as PDF</span>
              </span>
            </DropdownMenuItem>
            {!entry.appId && (
              <DropdownMenuItem onSelect={() => void withToast("Downloading the .pbix…", downloadPbix(entry), "Downloaded.")}>
                <FileDownIcon />
                <span className="flex flex-col">
                  Power BI file (.pbix)
                  <span className="text-[11px] text-muted-foreground">Connected to the online model</span>
                </span>
              </DropdownMenuItem>
            )}
          </DropdownMenuContent>
        </DropdownMenu>
      )}
    </>
  );
}
