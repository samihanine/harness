/**
 * Every AI conversation of this browser in one table (Settings): its screen, the item it is about
 * (report, research, deck…), its size and its ratings; searched and filtered, each downloadable (or all
 * of them), e.g. to review the bad answers.
 */
import { useMemo, useState } from "react";
import { DownloadIcon, ThumbsDownIcon, ThumbsUpIcon, Trash2Icon } from "lucide-react";
import { deleteConversation, downloadConversation } from "./chat";
import { Badge } from "./ui/badge";
import { Button } from "./ui/button";
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle } from "./ui/dialog";
import { Input } from "./ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "./ui/table";
import { SCREEN_LABELS, type Screen } from "@/config";
import { useCollection, type Conversation } from "@/lib/store";

/** The screen of a conversation's scope ("research:<id>" → research). */
const screenOf = (scope: string) => scope.split(":")[0] as Screen;

const FILTERS = { all: "All", rated: "Rated", good: "Good", bad: "Not good" } as const;

export function ConversationsDialog({ open, onOpenChange }: { open: boolean; onOpenChange: (open: boolean) => void }) {
  const conversations = useCollection("conversations");
  const reports = useCollection("reports");
  const datasets = useCollection("datasets");
  const researches = useCollection("researches");
  const decks = useCollection("decks");
  const schemas = useCollection("schemas");
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<keyof typeof FILTERS>("all");
  /** The name of the item a conversation is about, from its scope. */
  const itemOf = (scope: string) => {
    const [screen, id] = scope.split(":");
    const lists: Record<string, { id: string; name?: string; title?: string }[]> = { viewer: reports, editor: reports, reports, datasets, research: researches, slides: decks, excel: schemas };
    const item = lists[screen]?.find((x) => x.id === id);
    return item ? (item.name ?? item.title ?? id) : id && id !== "undefined" ? `${id.slice(0, 8)}… (removed)` : "—";
  };
  const rows = useMemo(
    () =>
      conversations
        .map((c) => {
          const answers = Object.values(c.ratings ?? {});
          const all = [...answers, ...(c.rating ? [c.rating] : [])];
          return { c, tab: SCREEN_LABELS[screenOf(c.scope)] ?? screenOf(c.scope), item: itemOf(c.scope), good: all.filter((r) => r.score === "good").length, bad: all.filter((r) => r.score === "bad").length, notes: all.map((r) => r.note).filter(Boolean).join(" · ") };
        })
        .filter((r) => (filter === "all" ? true : filter === "rated" ? r.good + r.bad > 0 : filter === "good" ? r.good > 0 : r.bad > 0))
        .filter((r) => !search.trim() || `${r.c.title} ${r.tab} ${r.item} ${r.notes}`.toLowerCase().includes(search.trim().toLowerCase()))
        .sort((a, b) => b.c.updatedAt - a.c.updatedAt),
    [conversations, reports, datasets, researches, decks, schemas, search, filter], // eslint-disable-line react-hooks/exhaustive-deps
  );
  const downloadAll = (list: Conversation[]) => {
    const json = JSON.stringify({ exportedAt: new Date().toISOString(), conversations: list }, null, 2);
    const url = URL.createObjectURL(new Blob([json], { type: "application/json" }));
    Object.assign(document.createElement("a"), { href: url, download: `conversations-${new Date().toISOString().slice(0, 10)}.json` }).click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  };
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="flex max-h-[85vh] flex-col gap-3 sm:max-w-5xl">
        <DialogHeader>
          <DialogTitle>Conversations</DialogTitle>
          <DialogDescription>Every conversation with the assistant in this browser, with the screen and the item it is about, and its ratings.</DialogDescription>
        </DialogHeader>
        <div className="flex flex-wrap items-center gap-2">
          <Input className="h-9 max-w-xs" placeholder="Search (title, screen, item, notes)…" value={search} onChange={(e) => setSearch(e.target.value)} />
          {(Object.keys(FILTERS) as (keyof typeof FILTERS)[]).map((k) => (
            <Button key={k} size="sm" variant={filter === k ? "secondary" : "ghost"} className="h-8 rounded-full" onClick={() => setFilter(k)}>
              {FILTERS[k]}
            </Button>
          ))}
          <span className="ml-auto text-[12px] text-muted-foreground">{rows.length} conversation(s)</span>
          <Button size="sm" variant="outline" disabled={!rows.length} onClick={() => downloadAll(rows.map((r) => r.c))}>
            <DownloadIcon /> Download these
          </Button>
        </div>
        <div className="min-h-0 flex-1 overflow-auto rounded-lg border">
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>Conversation</TableHead>
                <TableHead>Screen</TableHead>
                <TableHead>About</TableHead>
                <TableHead className="text-right">Messages</TableHead>
                <TableHead>Ratings</TableHead>
                <TableHead>Updated</TableHead>
                <TableHead />
              </TableRow>
            </TableHeader>
            <TableBody>
              {rows.map(({ c, tab, item, good, bad, notes }) => (
                <TableRow key={c.id}>
                  <TableCell className="max-w-72 truncate font-medium" title={c.title}>
                    {c.title}
                  </TableCell>
                  <TableCell>
                    <Badge variant="secondary">{tab}</Badge>
                  </TableCell>
                  <TableCell className="max-w-56 truncate" title={item}>
                    {item}
                  </TableCell>
                  <TableCell className="text-right tabular-nums">{c.messages?.length ?? 0}</TableCell>
                  <TableCell title={notes || undefined}>
                    <span className="flex items-center gap-2 text-[12.5px]">
                      {good > 0 && (
                        <span className="flex items-center gap-1 text-success">
                          <ThumbsUpIcon className="size-3.5" /> {good}
                        </span>
                      )}
                      {bad > 0 && (
                        <span className="flex items-center gap-1 text-destructive">
                          <ThumbsDownIcon className="size-3.5" /> {bad}
                        </span>
                      )}
                      {!good && !bad && <span className="text-muted-foreground">—</span>}
                      {notes && <span className="max-w-40 truncate text-muted-foreground">{notes}</span>}
                    </span>
                  </TableCell>
                  <TableCell className="whitespace-nowrap text-muted-foreground">{new Date(c.updatedAt).toLocaleString()}</TableCell>
                  <TableCell>
                    <span className="flex justify-end gap-0.5">
                      <button type="button" className="icon-btn size-7" aria-label="Download the conversation" onClick={() => downloadConversation(c)}>
                        <DownloadIcon className="size-3.5" />
                      </button>
                      <button type="button" className="icon-btn size-7 hover:text-destructive" aria-label="Delete the conversation" onClick={() => confirm(`Delete "${c.title}"?`) && void deleteConversation(c)}>
                        <Trash2Icon className="size-3.5" />
                      </button>
                    </span>
                  </TableCell>
                </TableRow>
              ))}
              {!rows.length && (
                <TableRow>
                  <TableCell colSpan={7} className="py-10 text-center text-muted-foreground">
                    No conversation.
                  </TableCell>
                </TableRow>
              )}
            </TableBody>
          </Table>
        </div>
      </DialogContent>
    </Dialog>
  );
}
