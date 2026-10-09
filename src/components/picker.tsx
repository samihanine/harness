/** Combobox (shadcn style): searchable list of library items, with an "Add" entry that takes a link. */
import { useState } from "react";
import { CheckIcon, ChevronsUpDownIcon, PlusIcon, SearchIcon } from "lucide-react";
import { errorText } from "@/lib/ms";
import { Popover } from "./popover";

type Item = { id: string; label: string; hint?: string };

export function Picker({
  items,
  value,
  onChange,
  placeholder,
  add,
  className = "w-60",
}: {
  items: Item[];
  value?: string;
  onChange: (id: string) => void;
  placeholder: string;
  /** Adds an item from a link and returns its id. */
  add?: { label: string; placeholder: string; run: (link: string) => Promise<string> };
  className?: string;
}) {
  const [search, setSearch] = useState("");
  const [adding, setAdding] = useState<{ link: string; busy?: boolean; error?: string }>();
  const current = items.find((i) => i.id === value);
  const shown = items.filter((i) => `${i.label} ${i.hint ?? ""}`.toLowerCase().includes(search.toLowerCase()));
  const submit = async (close: () => void) => {
    if (!add || !adding?.link.trim()) return;
    setAdding({ ...adding, busy: true, error: undefined });
    try {
      onChange(await add.run(adding.link.trim()));
      setAdding(undefined);
      close();
    } catch (e) {
      setAdding({ ...adding, busy: false, error: errorText(e) });
    }
  };
  return (
    <span className={className}>
      <Popover
        full
        width={320}
        trigger={(open) => (
          <button type="button" onClick={() => (setSearch(""), setAdding(undefined), open())} className="flex h-8 w-full items-center gap-2 rounded-md border bg-card px-2.5 text-left text-[13px] shadow-xs hover:bg-muted/50">
            <span className={`flex-1 truncate ${current ? "" : "text-muted-foreground"}`}>{current?.label ?? placeholder}</span>
            <ChevronsUpDownIcon className="size-3.5 shrink-0 text-muted-foreground" />
          </button>
        )}
      >
        {(close) =>
          adding ? (
            <div className="flex flex-col gap-2 p-1">
              <span className="label">{add!.label}</span>
              <input autoFocus className="input" placeholder={add!.placeholder} value={adding.link} onChange={(e) => setAdding({ ...adding, link: e.target.value })} onKeyDown={(e) => e.key === "Enter" && void submit(close)} />
              {adding.error && <p className="text-[11px] text-destructive">{adding.error}</p>}
              <div className="flex justify-end gap-1.5">
                <button type="button" className="btn h-7" onClick={() => setAdding(undefined)}>
                  Back
                </button>
                <button type="button" className="btn-primary h-7" disabled={adding.busy || !adding.link.trim()} onClick={() => void submit(close)}>
                  {adding.busy ? "Reading…" : "Add"}
                </button>
              </div>
            </div>
          ) : (
            <>
              <label className="flex items-center gap-2 border-b px-1 pb-1.5">
                <SearchIcon className="size-3.5 text-muted-foreground" />
                <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search…" className="h-7 flex-1 bg-transparent outline-none" />
              </label>
              <div className="max-h-72 overflow-y-auto">
                {shown.map((i) => (
                  <button key={i.id} type="button" onClick={() => (onChange(i.id), close())} className="flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left hover:bg-muted">
                    <CheckIcon className={`size-3.5 shrink-0 ${i.id === value ? "" : "invisible"}`} />
                    <span className="min-w-0 flex-1 truncate">{i.label}</span>
                    {i.hint && <span className="shrink-0 text-[11px] text-muted-foreground">{i.hint}</span>}
                  </button>
                ))}
                {shown.length === 0 && <p className="px-2 py-2 text-[12px] text-muted-foreground">Nothing found.</p>}
              </div>
              {add && (
                <button type="button" onClick={() => setAdding({ link: "" })} className="flex items-center gap-2 border-t px-2 pt-2 text-left text-[13px] hover:text-foreground">
                  <PlusIcon className="size-3.5" /> {add.label}
                </button>
              )}
            </>
          )
        }
      </Popover>
    </span>
  );
}
