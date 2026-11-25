/**
 * Choosing the item a screen works on: the screen's own list (kept in this browser, renamable), and "Add"
 * to pick a report or a semantic model the user can open in Power BI — workspaces and apps, searched page
 * by page, or a pasted link (an app's report or model too). The Builder lists only the reports it can
 * change: their rights are checked in the background for the page shown.
 */
import { useEffect, useMemo, useState } from "react";
import { CheckIcon, ChevronDownIcon, ChevronLeftIcon, ChevronRightIcon, Loader2Icon, PencilIcon, PlusIcon, SearchIcon, Trash2Icon } from "lucide-react";
import { toast } from "sonner";
import type { DatasetRef, ReportAccess, ReportRef } from "@shared/types";
import { Button } from "./ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "./ui/dialog";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger } from "./ui/dropdown-menu";
import { errorText, get, post } from "@/lib/api";
import { prepareDataset, prepareReport } from "@/lib/items";
import { item, remove, save, useItems } from "@/lib/store";
import { useItemParam } from "@/lib/params";
import { cn } from "@/lib/utils";

type CatalogItem = { id: string; name: string; groupId?: string; appId?: string; workspace: string; datasetId?: string };

const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
const PAGE = 20;

/**
 * Any Power BI link → ids: the browser's address of a report or a model (a workspace's, an app's, "My
 * workspace", shared), an embed link (reportEmbed?reportId=…&groupId=…), or a bare id.
 */
export function parseLink(url: string) {
  const g = (re: RegExp) => url.match(re)?.[1]?.toLowerCase();
  const report = g(/(?:reports|rdlreports)\/([0-9a-f-]{36})/i) ?? g(/[?&]reportId=([0-9a-f-]{36})/i);
  const dataset = g(/(?:datasets|semanticmodels)\/([0-9a-f-]{36})/i) ?? g(/[?&]datasetId=([0-9a-f-]{36})/i);
  const appId = g(/apps\/([0-9a-f-]{36})/i) ?? g(/[?&]appId=([0-9a-f-]{36})/i);
  const groupId = g(/groups\/([0-9a-f-]{36})/i) ?? g(/[?&]groupId=([0-9a-f-]{36})/i);
  const bare = !report && !dataset && !appId && !groupId ? url.match(GUID)?.[0]?.toLowerCase() : undefined;
  return report || dataset || bare ? { report, dataset, bare, appId, groupId } : undefined;
}
type Link = NonNullable<ReturnType<typeof parseLink>>;

/** What a link adds: a report (by its id), or a model — its own link, or the model of a report's link. */
async function addFromLink(kind: "report" | "dataset", l: Link) {
  if (kind === "report") {
    const id = l.report ?? l.bare;
    if (!id) throw new Error("This is the link of a semantic model, not of a report.");
    return (await addReport({ id, groupId: l.groupId, appId: l.appId })).id;
  }
  if (l.dataset || (l.bare && !l.report)) return (await addDataset({ id: (l.dataset ?? l.bare)!, groupId: l.groupId })).id;
  const q = new URLSearchParams({ id: l.report!, ...(l.groupId ? { groupId: l.groupId } : {}), ...(l.appId ? { appId: l.appId } : {}) });
  const r = await get<ReportRef>(`/powerbi/report?${q}`);
  if (!r.datasetId) throw new Error(`The model of "${r.name}" cannot be read with your rights.`);
  return (await addDataset({ id: r.datasetId, groupId: r.datasetGroupId })).id;
}

/** Reads a report's details from the server (model, access), keeps it, and reads its structure in the background. */
export async function addReport(ref: { id: string; groupId?: string; appId?: string }) {
  const q = new URLSearchParams({ id: ref.id, ...(ref.groupId ? { groupId: ref.groupId } : {}), ...(ref.appId ? { appId: ref.appId } : {}) });
  const r = await get<ReportRef>(`/powerbi/report?${q}`);
  const kept = await item("report", r.id);
  const saved = { ...r, ...(kept ? { name: kept.name, structure: kept.structure, structureAt: kept.structureAt } : {}), addedAt: kept?.addedAt ?? Date.now() };
  await save("report", saved);
  prepareReport(saved);
  return saved;
}

export async function addDataset(ref: { id: string; groupId?: string; name?: string }) {
  const q = new URLSearchParams({ id: ref.id, ...(ref.groupId ? { groupId: ref.groupId } : {}) });
  const d = ref.name ? { id: ref.id, groupId: ref.groupId, name: ref.name } : await get<DatasetRef>(`/powerbi/dataset?${q}`);
  const kept = await item("dataset", d.id);
  const saved = { ...d, ...(kept ? { name: kept.name } : {}), addedAt: kept?.addedAt ?? Date.now() };
  await save("dataset", saved);
  prepareDataset(saved);
  return saved;
}

export function AddDialog({ kind, open, onOpenChange, onAdded, editable }: { kind: "report" | "dataset"; open: boolean; onOpenChange: (o: boolean) => void; onAdded: (id: string) => void; editable?: boolean }) {
  const [catalog, setCatalog] = useState<CatalogItem[]>();
  const [error, setError] = useState<string>();
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(0);
  const [busy, setBusy] = useState<string>();
  const [access, setAccess] = useState<Record<string, ReportAccess>>({});
  const [showAll, setShowAll] = useState(false);
  useEffect(() => {
    if (!open || catalog) return;
    get<CatalogItem[]>(kind === "report" ? "/powerbi/reports" : "/powerbi/datasets").then(setCatalog, (e) => setError(errorText(e)));
  }, [open, kind, catalog]);
  useEffect(() => setPage(0), [query, showAll]);
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  const matching = (catalog ?? []).filter((c) => words.every((w) => `${c.name} ${c.workspace}`.toLowerCase().includes(w)));
  // The Builder: apps' reports are never editable; the others are checked for the page shown.
  const candidates = editable && !showAll ? matching.filter((c) => !c.appId && access[c.id] !== "view") : matching;
  const shown = candidates.slice(page * PAGE, page * PAGE + PAGE);
  const unknown = editable ? shown.filter((c) => !c.appId && !access[c.id]) : [];
  useEffect(() => {
    if (!unknown.length) return;
    let gone = false;
    post<Record<string, ReportAccess>>("/powerbi/access", unknown.map((c) => ({ id: c.id, groupId: c.groupId }))).then((a) => !gone && setAccess((x) => ({ ...x, ...a })), () => undefined);
    return () => void (gone = true);
  }, [unknown.map((c) => c.id).join()]); // eslint-disable-line react-hooks/exhaustive-deps
  const choose = async (ref: { id: string; groupId?: string; appId?: string; name?: string } | Link) => {
    setBusy("id" in ref ? ref.id : "link");
    try {
      const id = !("id" in ref) ? await addFromLink(kind, ref) : kind === "report" ? (await addReport(ref)).id : (await addDataset(ref)).id;
      onAdded(id);
      onOpenChange(false);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  const link = query.includes("powerbi.com") || GUID.test(query) ? parseLink(query) : undefined;
  const pages = Math.max(1, Math.ceil(candidates.length / PAGE));
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-xl rounded-3xl">
        <DialogHeader>
          <DialogTitle>{kind === "report" ? (editable ? "Add a report you can edit" : "Add a report") : "Add a semantic model"}</DialogTitle>
        </DialogHeader>
        <div className="flex items-center gap-2 rounded-xl border px-3">
          <SearchIcon className="size-4 text-muted-foreground" />
          <input autoFocus value={query} onChange={(e) => setQuery(e.target.value)} onKeyDown={(e) => e.key === "Enter" && link && !busy && void choose(link)} placeholder={kind === "report" ? "Search, or paste the report's address from your browser" : "Search, or paste the link of a model or of one of its reports"} className="h-10 flex-1 bg-transparent text-[14px] outline-none" />
        </div>
        {link && (
          <Button variant="dark" onClick={() => choose(link)} disabled={!!busy}>
            {busy === "link" ? <Loader2Icon className="animate-spin" /> : null} {kind === "dataset" && link.report && !link.dataset ? "Add this report's semantic model" : "Add from the link"}
          </Button>
        )}
        <div className="max-h-[50vh] min-h-40 overflow-y-auto">
          {error && <div className="p-3 text-[13px] text-destructive">{error}</div>}
          {!catalog && !error && (
            <div className="flex items-center gap-2 p-3 text-[13px] text-muted-foreground">
              <Loader2Icon className="size-4 animate-spin" /> Loading your {kind === "report" ? "reports" : "models"} (workspaces and apps)…
            </div>
          )}
          {shown.map((c) => {
            const a = c.appId ? "view" : access[c.id];
            return (
              <button key={c.id} onClick={() => choose(c)} disabled={!!busy || (editable && !showAll && !a)} className="flex w-full items-center gap-3 rounded-xl px-3 py-2 text-left hover:bg-muted disabled:opacity-60">
                <div className="min-w-0 flex-1">
                  <div className="truncate text-[14px] font-medium">{c.name}</div>
                  <div className="truncate text-[12px] text-muted-foreground">{c.workspace}</div>
                </div>
                {editable && (a ? <span className="badge bg-muted text-muted-foreground">{a === "edit" ? "editable" : a === "editor" ? "editor mode" : "view only"}</span> : <Loader2Icon className="size-3.5 animate-spin text-muted-foreground" />)}
                {busy === c.id && <Loader2Icon className="size-4 animate-spin" />}
              </button>
            );
          })}
          {catalog && !shown.length && <div className="p-3 text-[13px] text-muted-foreground">{editable && !showAll ? "No editable report on this page." : "Nothing matches."}</div>}
        </div>
        <div className="flex items-center gap-2 text-[12.5px] text-muted-foreground">
          {editable && (
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={showAll} onChange={(e) => setShowAll(e.target.checked)} /> Show view-only reports too
            </label>
          )}
          <div className="flex-1" />
          <Button variant="ghost" size="icon-sm" disabled={page === 0} onClick={() => setPage(page - 1)} aria-label="Previous page"><ChevronLeftIcon /></Button>
          <span className="tabular-nums">{page + 1} / {pages}</span>
          <Button variant="ghost" size="icon-sm" disabled={page >= pages - 1} onClick={() => setPage(page + 1)} aria-label="Next page"><ChevronRightIcon /></Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}

/** The chosen report or model of a screen, with its list and "Add". */
export function usePicked<K extends "report" | "dataset">(kind: K, screen: string) {
  const items = useItems(kind);
  // In the address (each browser tab its own), else the last one chosen here.
  const [id, setId] = useItemParam(`picked:${screen}:${kind}`);
  const sorted = useMemo(() => [...(items ?? [])].sort((a, b) => a.name.localeCompare(b.name)), [items]);
  const current = sorted.find((x) => x.id === id) ?? (items && !id ? sorted[0] : undefined);
  return { items: sorted, current, select: setId, loading: !items };
}

/** The picker in a screen's header: the item's name, a menu to switch, rename, remove or add. */
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
                const next = prompt("Name (in this app only)", name(it));
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
                if (confirm(`Remove "${name(it)}" from the list?`)) onDelete(it.id);
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

export function ItemPicker<T extends { id: string; name: string }>({ kind, items, current, onSelect, label, extra, editable }: { kind: "report" | "dataset"; items: T[]; current?: T; onSelect: (id: string) => void; label?: string; extra?: (item: T) => string | undefined; editable?: boolean }) {
  const [adding, setAdding] = useState(false);
  return (
    <>
      <PickerMenu
        items={items}
        current={current}
        name={(t) => t.name}
        onSelect={onSelect}
        onRename={(t, name) => void save(kind, { ...(t as object), name } as never)}
        onDelete={(id) => void remove(kind, id)}
        addLabel={kind === "report" ? "Add a report…" : "Add a semantic model…"}
        onAdd={() => setAdding(true)}
        empty={label ?? (kind === "report" ? "Choose a report" : "Choose a model")}
        extra={extra}
      />
      <AddDialog kind={kind} open={adding} onOpenChange={setAdding} onAdded={onSelect} editable={editable} />
    </>
  );
}

/** "Add a report / model" as a button (empty states). */
export function AddButton({ kind, onAdded, editable }: { kind: "report" | "dataset"; onAdded: (id: string) => void; editable?: boolean }) {
  const [open, setOpen] = useState(false);
  return (
    <>
      <Button variant="dark" size="sm" className="mt-3" onClick={() => setOpen(true)}>
        <PlusIcon /> {kind === "report" ? "Add a report" : "Add a semantic model"}
      </Button>
      <AddDialog kind={kind} open={open} onOpenChange={setOpen} onAdded={onAdded} editable={editable} />
    </>
  );
}
