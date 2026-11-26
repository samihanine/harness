/**
 * Choosing the reports and models a screen works on, from the team's catalog (Confluence, see
 * src/lib/library.ts): the screen's picker (items hidden by this user left out, "Manage the list" to Settings
 * when the catalog is editable), checklists of items (Data Scientist, Presenter, Builder references), and the
 * dialog that adds an item to the catalog — any report or model the user can open in Power BI, found in their
 * workspaces and apps or by a pasted link.
 */
import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "@tanstack/react-router";
import { CheckIcon, ChevronDownIcon, ChevronLeftIcon, ChevronRightIcon, EyeIcon, EyeOffIcon, Loader2Icon, PencilIcon, PlusIcon, SearchIcon, Settings2Icon, Trash2Icon } from "lucide-react";
import { toast } from "sonner";
import type { LibraryEntry, LibraryKind, ReportRef } from "@shared/types";
import { datasetUrl, isPowerBiLink, parseLink, reportUrl } from "@shared/links";
import { Button } from "./ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "./ui/dialog";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger } from "./ui/dropdown-menu";
import { errorText, get } from "@/lib/api";
import { openReport, useHidden, useLibrary } from "@/lib/library";
import { useItemParam } from "@/lib/params";
import { useSettings } from "@/lib/settings";
import { useItems } from "@/lib/store";
import { cn } from "@/lib/utils";

type CatalogItem = { id: string; name: string; groupId?: string; appId?: string; workspace: string };
const PAGE = 20;
const noun = (kind: LibraryKind) => (kind === "report" ? "report" : "semantic model");

/** The catalog's items this user did not hide (all with `all`). */
export function useVisible(kind: LibraryKind, all = false) {
  const { entries, error } = useLibrary(kind);
  const [hidden, setHidden] = useHidden(kind);
  const items = useMemo(() => (entries ?? []).filter((e) => e.id && (all || !hidden.includes(e.id))), [entries, hidden, all]);
  return { items, entries, error, hidden, setHidden, loading: !entries && !error };
}

/**
 * The report of a screen: chosen from the catalog (in the address, else the last one chosen here), opened on
 * first use (its details read: refused when the user cannot open it).
 */
export function usePicked(screen: string) {
  const { items, entries, error: listError, loading } = useVisible("report");
  const [id, setId] = useItemParam(`picked:${screen}:report`);
  const entry = items.find((x) => x.id === id) ?? (entries ?? []).find((x) => x.id === id) ?? (!id ? items[0] : undefined);
  const stored = useItems("report")?.find((r) => r.id === entry?.id);
  const [error, setError] = useState<string>();
  useEffect(() => {
    setError(undefined);
    if (entry && !stored) openReport(entry).catch((e) => setError(`You cannot open "${entry.name}": ${errorText(e)}`));
  }, [entry?.id, !!stored]); // eslint-disable-line react-hooks/exhaustive-deps
  const current: ReportRef | undefined = entry && stored ? { ...stored, name: entry.name } : undefined;
  return { items, entry, current, select: setId, loading: loading || (!!entry && !stored && !error), error: error ?? listError };
}

/** "Manage the list" (Settings), shown when the catalog can be changed here. */
function useManage(kind: LibraryKind) {
  const settings = useSettings();
  const navigate = useNavigate();
  return settings?.confluence.catalogEdit ? () => void navigate({ to: "/settings", search: { tab: kind === "report" ? "reports" : "models" } as never }) : undefined;
}

/** The picker in a screen's header: the catalog's reports, hide one, show the hidden ones, manage the list. */
export function ItemPicker({ kind, current, onSelect }: { kind: LibraryKind; current?: { id?: string; name: string }; onSelect: (id: string) => void }) {
  const [showHidden, setShowHidden] = useState(false);
  const { items, hidden, setHidden } = useVisible(kind, showHidden);
  const manage = useManage(kind);
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button className="flex max-w-[420px] items-center gap-1 rounded-full px-2 py-1 text-[15px] text-muted-foreground hover:bg-muted hover:text-foreground">
          <span className="truncate">{current?.name ?? `Choose a ${noun(kind)}`}</span>
          <ChevronDownIcon className="size-4 shrink-0" />
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="max-h-[70vh] w-96 overflow-y-auto">
        {items.map((it) => {
          const off = hidden.includes(it.id!);
          return (
            <DropdownMenuItem key={it.id} onSelect={() => onSelect(it.id!)} className="group items-start">
              <CheckIcon className={cn("mt-0.5 size-4", it.id === current?.id ? "opacity-100" : "opacity-0")} />
              <div className={cn("min-w-0 flex-1", off && "opacity-50")}>
                <div className="truncate">{it.name}</div>
                {it.description && <div className="line-clamp-2 text-[11.5px] text-muted-foreground">{it.description}</div>}
              </div>
              <button className="opacity-0 group-hover:opacity-100" title={off ? "Show it in my list" : "Hide it from my list"} onClick={(e) => (e.stopPropagation(), setHidden(off ? hidden.filter((x) => x !== it.id) : [...hidden, it.id!]))}>
                {off ? <EyeIcon className="size-3.5" /> : <EyeOffIcon className="size-3.5" />}
              </button>
            </DropdownMenuItem>
          );
        })}
        {!items.length && <div className="px-2 py-1.5 text-[13px] text-muted-foreground">No {noun(kind)} in the catalog{hidden.length ? " (some are hidden)" : ""}.</div>}
        <DropdownMenuSeparator />
        {hidden.length > 0 && (
          <DropdownMenuItem onSelect={(e) => (e.preventDefault(), setShowHidden(!showHidden))}>
            {showHidden ? <EyeOffIcon /> : <EyeIcon />} {showHidden ? "Leave out the hidden ones" : `Show the hidden ones (${hidden.length})`}
          </DropdownMenuItem>
        )}
        {manage && (
          <DropdownMenuItem onSelect={manage}>
            <Settings2Icon /> Manage the list…
          </DropdownMenuItem>
        )}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

/** Items of the catalog to check (the ones a document may use), with "Manage the list". */
export function CatalogChecklist({ kind, ids, setIds, title, hint }: { kind: LibraryKind; ids: string[]; setIds: (ids: string[]) => void; title: string; hint: string }) {
  const { items, loading, error } = useVisible(kind);
  const manage = useManage(kind);
  return (
    <div className="mx-auto flex max-w-3xl flex-col gap-3 p-6">
      <div className="flex items-center gap-3">
        <div className="flex-1">
          <div className="text-[15px] font-semibold">{title}</div>
          <div className="text-[13px] text-muted-foreground">{hint}</div>
        </div>
        {manage && <Button variant="outline" size="sm" onClick={manage}><Settings2Icon /> Manage the list</Button>}
      </div>
      {loading && <div className="flex items-center gap-2 text-[13px] text-muted-foreground"><Loader2Icon className="size-4 animate-spin" /> Reading the catalog…</div>}
      {error && <div className="text-[13px] text-destructive">{error}</div>}
      {items.map((r) => {
        const on = ids.includes(r.id!);
        return (
          <button key={r.id} onClick={() => setIds(on ? ids.filter((x) => x !== r.id) : [...ids, r.id!])} className={cn("flex items-start gap-3 rounded-2xl border px-4 py-3 text-left", on && "border-foreground/40 bg-muted/50")}>
            <span className={cn("mt-0.5 flex size-5 shrink-0 items-center justify-center rounded-md border", on && "border-foreground bg-foreground text-background")}>{on && <CheckIcon className="size-3.5" />}</span>
            <div className="min-w-0 flex-1">
              <div className="text-[14px]">{r.name}</div>
              {r.description && <div className="line-clamp-2 text-[12.5px] text-muted-foreground">{r.description}</div>}
            </div>
          </button>
        );
      })}
      {!loading && !error && !items.length && <div className="text-[13px] text-muted-foreground">The catalog has no {noun(kind)} yet.</div>}
    </div>
  );
}

/** Adds a report or a model to the catalog: from the user's workspaces and apps, page by page, or a pasted link. */
export function AddDialog({ kind, open, onOpenChange, onAdd }: { kind: LibraryKind; open: boolean; onOpenChange: (o: boolean) => void; onAdd: (link: string) => Promise<void> }) {
  const [catalog, setCatalog] = useState<CatalogItem[]>();
  const [error, setError] = useState<string>();
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(0);
  const [busy, setBusy] = useState<string>();
  useEffect(() => {
    if (!open || catalog) return;
    get<CatalogItem[]>(kind === "report" ? "/powerbi/reports" : "/powerbi/datasets").then(setCatalog, (e) => setError(errorText(e)));
  }, [open, kind, catalog]);
  useEffect(() => setPage(0), [query]);
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  const matching = (catalog ?? []).filter((c) => words.every((w) => `${c.name} ${c.workspace}`.toLowerCase().includes(w)));
  const shown = matching.slice(page * PAGE, page * PAGE + PAGE);
  const pages = Math.max(1, Math.ceil(matching.length / PAGE));
  const link = isPowerBiLink(query) ? parseLink(query) : undefined;
  const add = async (key: string, url: string) => {
    setBusy(key);
    try {
      await onAdd(url);
      onOpenChange(false);
      setQuery("");
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  const urlOf = (c: CatalogItem) => (kind === "report" ? reportUrl(c) : datasetUrl(c));
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-xl rounded-3xl">
        <DialogHeader>
          <DialogTitle>Add a {noun(kind)} to the catalog</DialogTitle>
        </DialogHeader>
        <div className="flex items-center gap-2 rounded-xl border px-3">
          <SearchIcon className="size-4 text-muted-foreground" />
          <input autoFocus value={query} onChange={(e) => setQuery(e.target.value)} onKeyDown={(e) => e.key === "Enter" && link && !busy && void add("link", query.trim())} placeholder={kind === "report" ? "Search, or paste the report's address from your browser" : "Search, or paste the link of a model or of one of its reports"} className="h-10 flex-1 bg-transparent text-[14px] outline-none" />
        </div>
        {link && (
          <Button variant="dark" onClick={() => void add("link", query.trim())} disabled={!!busy}>
            {busy === "link" ? <Loader2Icon className="animate-spin" /> : <PlusIcon />} {kind === "dataset" && link.report && !link.dataset ? "Add this report's semantic model" : "Add from the link"}
          </Button>
        )}
        <div className="max-h-[50vh] min-h-40 overflow-y-auto">
          {error && <div className="p-3 text-[13px] text-destructive">{error}</div>}
          {!catalog && !error && (
            <div className="flex items-center gap-2 p-3 text-[13px] text-muted-foreground">
              <Loader2Icon className="size-4 animate-spin" /> Loading your {kind === "report" ? "reports" : "models"} (workspaces and apps)…
            </div>
          )}
          {shown.map((c) => (
            <button key={c.id} onClick={() => void add(c.id, urlOf(c))} disabled={!!busy} className="flex w-full items-center gap-3 rounded-xl px-3 py-2 text-left hover:bg-muted disabled:opacity-60">
              <div className="min-w-0 flex-1">
                <div className="truncate text-[14px] font-medium">{c.name}</div>
                <div className="truncate text-[12px] text-muted-foreground">{c.workspace}</div>
              </div>
              {busy === c.id && <Loader2Icon className="size-4 animate-spin" />}
            </button>
          ))}
          {catalog && !shown.length && <div className="p-3 text-[13px] text-muted-foreground">Nothing matches.</div>}
        </div>
        <div className="flex items-center justify-end gap-2 text-[12.5px] text-muted-foreground">
          <Button variant="ghost" size="icon-sm" disabled={page === 0} onClick={() => setPage(page - 1)} aria-label="Previous page"><ChevronLeftIcon /></Button>
          <span className="tabular-nums">{page + 1} / {pages}</span>
          <Button variant="ghost" size="icon-sm" disabled={page >= pages - 1} onClick={() => setPage(page + 1)} aria-label="Next page"><ChevronRightIcon /></Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}

/** The menu of a screen's own documents (analyses, decks): switch, rename, remove, create. */
export function PickerMenu<T extends { id: string }>({ items, current, name, onSelect, onRename, onDelete, addLabel, onAdd, empty, extra }: { items: T[]; current?: T; name: (t: T) => string; onSelect: (id: string) => void; onRename: (t: T, name: string) => void; onDelete: (id: string) => void; addLabel: string; onAdd: () => void; empty: string; extra?: (t: T) => string | undefined }) {
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button className="flex max-w-[420px] items-center gap-1 rounded-full px-2 py-1 text-[15px] text-muted-foreground hover:bg-muted hover:text-foreground">
          <span className="truncate">{current ? name(current) : empty}</span>
          <ChevronDownIcon className="size-4 shrink-0" />
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="w-80">
        {items.map((it) => (
          <DropdownMenuItem key={it.id} onSelect={() => onSelect(it.id)} className="group">
            <CheckIcon className={cn("size-4", it.id === current?.id ? "opacity-100" : "opacity-0")} />
            <span className="min-w-0 flex-1 truncate">{name(it)}</span>
            {extra?.(it) && <span className="text-[11px] text-muted-foreground">{extra(it)}</span>}
            <button
              className="opacity-0 group-hover:opacity-100"
              onClick={(e) => {
                e.stopPropagation();
                const next = prompt("Name", name(it));
                if (next?.trim()) onRename(it, next.trim());
              }}
              aria-label="Rename"
            >
              <PencilIcon className="size-3.5" />
            </button>
            <button
              className="opacity-0 group-hover:opacity-100"
              onClick={(e) => {
                e.stopPropagation();
                if (confirm(`Remove "${name(it)}"?`)) onDelete(it.id);
              }}
              aria-label="Remove"
            >
              <Trash2Icon className="size-3.5" />
            </button>
          </DropdownMenuItem>
        ))}
        {items.length > 0 && <DropdownMenuSeparator />}
        <DropdownMenuItem onSelect={onAdd}>
          <PlusIcon /> {addLabel}
        </DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

export type { LibraryEntry };
