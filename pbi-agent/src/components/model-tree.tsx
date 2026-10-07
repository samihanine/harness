import { useState } from "react";
import type { ReactNode } from "react";
import { CalculatorIcon, ChevronRightIcon, DatabaseIcon, FolderIcon, HashIcon, SearchIcon, TableIcon, TypeIcon } from "lucide-react";
import type { Source } from "@/query/engine";
import { cn } from "@/lib/utils";

export type PickedField = { kind: "column" | "measure"; table: string; name: string; dataType?: string; ref: string };

const NUMERIC = /int|double|decimal|currency|number/i;

/** Source > table > (display folder) > columns and measures, with a search. */
export function ModelTree({
  sources,
  onPick,
  action,
}: {
  sources: Source[];
  onPick?: (source: Source, field: PickedField) => void;
  /** Rendered next to each field (e.g. an "add to…" menu). */
  action?: (source: Source, field: PickedField) => ReactNode;
}) {
  const [search, setSearch] = useState("");
  const query = search.trim().toLowerCase();
  const words = query.split(/[\s.›>\[\]']+/).filter(Boolean);
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <label className="mx-2 my-2 flex h-7 shrink-0 items-center gap-1.5 rounded-lg border bg-background px-2 focus-within:border-ring">
        <SearchIcon className="size-3.5 text-muted-foreground" />
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search table or field" className="min-w-0 flex-1 bg-transparent text-[12px] outline-none" />
      </label>
      <div className="min-h-0 flex-1 overflow-y-auto px-1 pb-2 text-[12px]">
        {sources.map((source) => (
          <Branch key={source.id} icon={<DatabaseIcon />} label={source.title} hint={source.language} open={sources.length === 1 || !!query} forceOpen={!!query}>
            {source.model.tables.map((table) => {
              const fields: (PickedField & { folder?: string })[] = [
                ...table.measures.map((m) => ({ kind: "measure" as const, table: table.name, name: m.name, ref: `[${m.name}]`, folder: m.displayFolder })),
                ...table.columns
                  .filter((c) => !c.hidden)
                  .map((c) => ({ kind: "column" as const, table: table.name, name: c.name, dataType: c.dataType, ref: `'${table.name}'[${c.name}]`, folder: c.displayFolder })),
              ].filter((f) => {
                // Every word must be found in the table, the folder or the field name ("sales amount", "Sales.Amount"…).
                const text = `${table.name} ${f.folder ?? ""} ${f.name}`.toLowerCase();
                return words.every((w) => text.includes(w));
              });
              if (fields.length === 0) return null;
              const folders = [...new Set(fields.map((f) => f.folder).filter(Boolean))] as string[];
              const leaf = (f: PickedField) => <Leaf key={f.ref} field={f} onPick={() => onPick?.(source, f)} action={action?.(source, f)} />;
              return (
                <Branch key={table.name} icon={<TableIcon />} label={table.name} hint={query ? String(fields.length) : undefined} open={!!query} forceOpen={!!query}>
                  {folders.map((folder) => (
                    <Branch key={folder} icon={<FolderIcon />} label={folder} open={!!query} forceOpen={!!query}>
                      {fields.filter((f) => f.folder === folder).map(leaf)}
                    </Branch>
                  ))}
                  {fields.filter((f) => !f.folder).map(leaf)}
                </Branch>
              );
            })}
          </Branch>
        ))}
      </div>
    </div>
  );
}

function Branch({ icon, label, hint, open: initial, forceOpen, children }: { icon: ReactNode; label: string; hint?: string; open: boolean; forceOpen?: boolean; children: ReactNode }) {
  const [open, setOpen] = useState(initial);
  const shown = forceOpen || open;
  return (
    <div>
      <button type="button" onClick={() => setOpen(!open)} className="flex w-full min-w-0 items-center gap-1 rounded-md px-1 py-1 text-left hover:bg-muted [&>svg]:size-3.5 [&>svg]:shrink-0 [&>svg]:text-muted-foreground">
        <ChevronRightIcon className={cn("transition-transform", shown && "rotate-90")} />
        {icon}
        <span className="truncate">{label}</span>
        {hint && <span className="ml-auto pl-1 text-[10px] text-muted-foreground">{hint}</span>}
      </button>
      {shown && <div className="ml-3 border-l pl-1">{children}</div>}
    </div>
  );
}

function Leaf({ field, onPick, action }: { field: PickedField; onPick: () => void; action?: ReactNode }) {
  const Icon = field.kind === "measure" ? CalculatorIcon : NUMERIC.test(field.dataType ?? "") ? HashIcon : TypeIcon;
  return (
    <div className="group flex min-w-0 items-center rounded-md hover:bg-muted">
      <button type="button" onClick={onPick} title={`${field.table} › ${field.name}`} className="flex min-w-0 flex-1 items-center gap-1.5 px-1.5 py-1 text-left">
        <Icon className={cn("size-3 shrink-0", field.kind === "measure" ? "text-amber-600 dark:text-amber-400" : "text-muted-foreground")} />
        <span className="truncate">{field.name}</span>
      </button>
      {action && <div className="opacity-0 group-hover:opacity-100">{action}</div>}
    </div>
  );
}
