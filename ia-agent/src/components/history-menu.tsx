import { ChevronDownIcon, MessageSquareIcon, Trash2Icon } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import type { Conversation } from "@/lib/types";

const ago = (time: number) => {
  const minutes = Math.round((Date.now() - time) / 60_000);
  if (minutes < 1) return "now";
  if (minutes < 60) return `${minutes}m`;
  if (minutes < 1440) return `${Math.round(minutes / 60)}h`;
  return `${Math.round(minutes / 1440)}d`;
};

export function HistoryMenu({
  title,
  list,
  currentId,
  onOpen,
  onRemove,
}: {
  title: string;
  list: Omit<Conversation, "items">[];
  currentId: string;
  onOpen: (id: string) => void;
  onRemove: (id: string) => void;
}) {
  return (
    <DropdownMenu>
      <DropdownMenuTrigger render={<Button variant="ghost" size="sm" className="min-w-0 shrink gap-1 px-2 font-medium" />}>
        <span className="truncate">{title}</span>
        <ChevronDownIcon className="size-3 shrink-0 text-muted-foreground" />
      </DropdownMenuTrigger>
      <DropdownMenuContent className="max-h-96 w-72">
        <DropdownMenuLabel>Conversations</DropdownMenuLabel>
        {list.length === 0 && <p className="px-2 py-1.5 text-[12px] text-muted-foreground">No conversation yet.</p>}
        {list.map((c) => (
          <DropdownMenuItem key={c.id} onClick={() => onOpen(c.id)} className="group/row gap-2">
            <MessageSquareIcon className={c.id === currentId ? "" : "opacity-40"} />
            <span className="truncate">{c.title}</span>
            <span className="ml-auto shrink-0 text-[11px] text-muted-foreground group-hover/row:hidden">{ago(c.updatedAt)}</span>
            <button
              type="button"
              aria-label="Delete"
              className="ml-auto hidden shrink-0 text-muted-foreground hover:text-destructive group-hover/row:block"
              onClick={(e) => {
                e.stopPropagation();
                onRemove(c.id);
              }}
            >
              <Trash2Icon className="size-3.5" />
            </button>
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}
