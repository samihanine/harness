/** The documents a screen keeps (analyses, decks, PowerPoints): choose, create, rename, delete. */
import { CheckIcon, ChevronDownIcon, PlusIcon, Trash2Icon } from "lucide-react";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger } from "./ui/dropdown-menu";
import { cn } from "@/lib/utils";

export function DocPicker<T extends { id: string }>({ items, current, name, onSelect, onCreate, onDelete, createLabel, empty }: { items: T[]; current?: T; name: (t: T) => string; onSelect: (id: string) => void; onCreate: () => void; onDelete: (id: string) => void; createLabel: string; empty: string }) {
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button className="flex max-w-[420px] items-center gap-1.5 rounded-full px-2 py-1 text-[26px] font-light tracking-tight hover:bg-muted">
          <span className="truncate">{current ? name(current) : empty}</span>
          <ChevronDownIcon className="size-5 shrink-0 text-muted-foreground" />
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="w-80">
        {items.map((it) => (
          <DropdownMenuItem key={it.id} onSelect={() => onSelect(it.id)} className="group">
            <CheckIcon className={cn("size-4", it.id === current?.id ? "opacity-100" : "opacity-0")} />
            <span className="min-w-0 flex-1 truncate">{name(it)}</span>
            <button
              className="opacity-0 group-hover:opacity-100"
              onClick={(e) => {
                e.stopPropagation();
                if (confirm(`Delete "${name(it)}"?`)) onDelete(it.id);
              }}
              aria-label="Delete"
            >
              <Trash2Icon className="size-3.5" />
            </button>
          </DropdownMenuItem>
        ))}
        {items.length > 0 && <DropdownMenuSeparator />}
        <DropdownMenuItem onSelect={onCreate}>
          <PlusIcon /> {createLabel}
        </DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}
