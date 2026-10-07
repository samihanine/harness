import { useEffect, useState } from "react";
import { CheckIcon, ChevronDownIcon, PlusIcon, SearchIcon } from "lucide-react";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { cn } from "@/lib/utils";

export type SearchOption = {
  value: string;
  label: string;
  /** Recency (timestamp or ISO date): the most recent option is selected by default. Without it, the last option wins. */
  recent?: number | string;
};

const time = (recent: SearchOption["recent"]) => (typeof recent === "string" ? Date.parse(recent) || 0 : (recent ?? 0));

/** The most recent option (by `recent`, else the last one). */
export const mostRecent = (options: SearchOption[]) =>
  options.reduce<SearchOption | undefined>((best, o) => (!best || time(o.recent) >= time(best.recent) ? o : best), undefined);

/**
 * Single-value select with a search box, an optional "add" button and a default value: when
 * nothing (valid) is selected, the most recent option is chosen.
 */
export function SearchSelect({
  value,
  onChange,
  options,
  placeholder = "Select…",
  onAdd,
  addLabel = "Add",
  className,
  size = "default",
  autoSelect = true,
}: {
  value: string | null;
  onChange: (value: string) => void;
  options: SearchOption[];
  placeholder?: string;
  onAdd?: () => void;
  addLabel?: string;
  className?: string;
  size?: "sm" | "default";
  /** Selects the most recent option when nothing is selected (default true). */
  autoSelect?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const selected = options.find((o) => o.value === value);
  const fallback = mostRecent(options)?.value;

  useEffect(() => {
    if (autoSelect && !selected && fallback) onChange(fallback);
  }, [selected, fallback]); // eslint-disable-line react-hooks/exhaustive-deps

  const words = search.toLowerCase().split(/[\s.›>\[\]']+/).filter(Boolean);
  const shown = options.filter((o) => words.every((w) => o.label.toLowerCase().includes(w)));
  return (
    <Popover
      open={open}
      onOpenChange={(o) => {
        setOpen(o);
        setSearch("");
      }}
    >
      <PopoverTrigger
        className={cn(
          "flex w-full items-center justify-between gap-2 rounded-lg border bg-transparent px-2.5 text-left outline-none transition-colors hover:bg-muted/50 focus-visible:border-ring focus-visible:ring-3 focus-visible:ring-ring/50",
          size === "sm" ? "h-7" : "h-8",
          className,
        )}
      >
        <span className={cn("truncate", !selected && "text-muted-foreground")}>{selected?.label ?? placeholder}</span>
        <ChevronDownIcon className="size-3.5 shrink-0 text-muted-foreground" />
      </PopoverTrigger>
      <PopoverContent align="start" className="w-(--anchor-width) min-w-56 gap-0 p-1">
        <div className="flex items-center gap-1.5 border-b px-2 pb-1">
          <SearchIcon className="size-3.5 text-muted-foreground" />
          <input
            autoFocus
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && shown[0]) {
                onChange(shown[0].value);
                setOpen(false);
              }
            }}
            placeholder="Search…"
            className="h-7 flex-1 bg-transparent outline-none"
          />
        </div>
        <div className="max-h-64 overflow-y-auto pt-1">
          {shown.map((o) => (
            <button
              key={o.value}
              type="button"
              onClick={() => {
                onChange(o.value);
                setOpen(false);
              }}
              className="flex w-full items-center gap-2 rounded-md px-2 py-1 text-left hover:bg-muted"
            >
              <span className="flex-1 truncate">{o.label}</span>
              {o.value === value && <CheckIcon className="size-3.5 shrink-0" />}
            </button>
          ))}
          {shown.length === 0 && <p className="px-2 py-1.5 text-[12px] text-muted-foreground">No result.</p>}
        </div>
        {onAdd && (
          <button
            type="button"
            onClick={() => {
              setOpen(false);
              onAdd();
            }}
            className="mt-1 flex w-full items-center gap-2 border-t px-2 pt-1.5 pb-1 text-left text-muted-foreground hover:text-foreground"
          >
            <PlusIcon className="size-3.5" /> {addLabel}
          </button>
        )}
      </PopoverContent>
    </Popover>
  );
}
