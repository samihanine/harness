/** Combobox (shadcn style): searchable list of library items, shown as the screen's title. */
import { useState, type ReactNode } from "react";
import { CheckIcon, ChevronsUpDownIcon, SearchIcon } from "lucide-react";
import { Popover } from "./popover";

export type PickerItem = { id: string; label: string; hint?: string; icon?: ReactNode };

export function Picker({
  items,
  value,
  onChange,
  placeholder,
  empty = "Nothing found.",
  footer,
  className = "",
}: {
  items: PickerItem[];
  value?: string;
  onChange: (id: string) => void;
  placeholder: string;
  /** Shown when the list is empty (e.g. where to add items). */
  empty?: ReactNode;
  /** Bottom of the list (e.g. a link to the library). */
  footer?: ReactNode;
  className?: string;
}) {
  const [search, setSearch] = useState("");
  const current = items.find((i) => i.id === value);
  const shown = items.filter((i) => `${i.label} ${i.hint ?? ""}`.toLowerCase().includes(search.toLowerCase()));
  return (
    <span className={`max-w-xs min-w-24 shrink-0 ${className}`}>
      <Popover
        full
        width={380}
        trigger={(open) => (
          <button type="button" onClick={() => (setSearch(""), open())} className="-ml-1 flex h-10 w-fit max-w-full items-center gap-2 rounded-lg px-2 text-left hover:bg-muted">
            {current?.icon}
            <span className={`truncate text-[15px] font-medium ${current ? "" : "text-muted-foreground"}`}>{current?.label ?? placeholder}</span>
            {current?.hint && <span className="shrink-0 rounded-md bg-muted px-1.5 py-0.5 text-[11.5px] font-medium text-muted-foreground">{current.hint}</span>}
            <ChevronsUpDownIcon className="size-4 shrink-0 text-muted-foreground" />
          </button>
        )}
      >
        {(close) => (
          <>
            <label className="flex items-center gap-2 border-b px-1.5 pb-2">
              <SearchIcon className="size-4 text-muted-foreground" />
              <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search…" className="h-8 flex-1 bg-transparent outline-none" />
            </label>
            <div className="max-h-80 overflow-y-auto">
              {shown.map((i) => (
                <button key={i.id} type="button" onClick={() => (onChange(i.id), close())} className="flex w-full items-center gap-2.5 rounded-lg px-2.5 py-2 text-left hover:bg-muted">
                  <CheckIcon className={`size-4 shrink-0 text-primary ${i.id === value ? "" : "invisible"}`} />
                  {i.icon}
                  <span className="min-w-0 flex-1 truncate">{i.label}</span>
                  {i.hint && <span className="shrink-0 text-[12px] text-muted-foreground">{i.hint}</span>}
                </button>
              ))}
              {shown.length === 0 && <div className="px-2.5 py-3 text-[13px] text-muted-foreground">{items.length ? "Nothing found." : empty}</div>}
            </div>
            {footer && <div className="border-t px-1 pt-2">{footer}</div>}
          </>
        )}
      </Popover>
    </span>
  );
}
