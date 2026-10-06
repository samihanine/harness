import { DownloadIcon } from "lucide-react";
import ExcelJS from "exceljs";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import type { QueryResult } from "@/pbi/api";
import { toCsv } from "@/query/engine";
import { IconButton } from "./icon-button";

const SHOWN = 500;

export const formatCell = (value: unknown) => {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return value.toLocaleString(undefined, { maximumFractionDigits: 4 });
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "string" && /^\d{4}-\d\d-\d\dT00:00:00/.test(value)) return value.slice(0, 10);
  return String(value);
};

function save(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  Object.assign(document.createElement("a"), { href: url, download: name }).click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function downloadResult(result: QueryResult, name: string, format: "csv" | "xlsx") {
  if (format === "csv") return save(new Blob([toCsv(result)], { type: "text/csv" }), `${name}.csv`);
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Result");
  sheet.addTable({
    name: "Result",
    ref: "A1",
    headerRow: true,
    style: { theme: "TableStyleLight1", showRowStripes: true },
    columns: result.columns.map((name) => ({ name, filterButton: true })),
    rows: result.rows.map((r) => result.columns.map((c) => (r[c] ?? null) as ExcelJS.CellValue)),
  });
  save(new Blob([await workbook.xlsx.writeBuffer()]), `${name}.xlsx`);
}

/** Query result: first rows shown, everything downloadable. */
export function ResultTable({ result, name = "result" }: { result: QueryResult; name?: string }) {
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <div className="flex h-8 shrink-0 items-center gap-2 px-3 text-[12px] text-muted-foreground">
        {result.rows.length.toLocaleString()} rows · {result.columns.length} columns
        {result.rows.length > SHOWN && ` · first ${SHOWN} shown`}
        <DropdownMenu>
          <DropdownMenuTrigger render={<IconButton label="Download" className="ml-auto" />}>
            <DownloadIcon />
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="w-36">
            <DropdownMenuItem onClick={() => void downloadResult(result, name, "xlsx")}>Excel (.xlsx)</DropdownMenuItem>
            <DropdownMenuItem onClick={() => void downloadResult(result, name, "csv")}>CSV</DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </div>
      <div className="min-h-0 flex-1 overflow-auto border-t">
        <table className="w-max min-w-full border-separate border-spacing-0 text-[12px]">
          <thead className="sticky top-0 bg-background">
            <tr>
              {result.columns.map((c) => (
                <th key={c} className="border-r border-b px-2 py-1.5 text-left font-medium whitespace-nowrap text-muted-foreground last:border-r-0">
                  {c}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {result.rows.slice(0, SHOWN).map((row, i) => (
              <tr key={i} className="hover:bg-muted/40">
                {result.columns.map((c) => (
                  <td key={c} className={`border-r border-b px-2 py-1 whitespace-nowrap last:border-r-0 ${typeof row[c] === "number" ? "text-right tabular-nums" : ""}`}>
                    {formatCell(row[c])}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
