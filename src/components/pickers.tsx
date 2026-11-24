/**
 * Choosing the item a screen works on: the screen's own list (kept in this browser), and "Add" to pick
 * a report or a semantic model the user can open in Power BI (searched, or a pasted link).
 */
import { useEffect, useMemo, useState } from "react";
import { CheckIcon, ChevronDownIcon, Loader2Icon, PlusIcon, SearchIcon, Trash2Icon } from "lucide-react";
import { toast } from "sonner";
import type { DatasetRef, ReportRef } from "@shared/types";
import { Button } from "./ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "./ui/dialog";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger } from "./ui/dropdown-menu";
import { errorText, get } from "@/lib/api";
import { remove, save, useItems, useLocal } from "@/lib/store";
import { cn } from "@/lib/utils";

type CatalogItem = { id: string; name: string; groupId?: string; workspace: string; datasetId?: string };

const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
/** app.powerbi.com links → ids. */
export function parseLink(url: string, kind: "report" | "dataset") {
  const id = url.match(kind === "report" ? /reports\/([0-9a-f-]{36})/i : /(?:datasets|semanticmodels)\/([0-9a-f-]{36})/i)?.[1] ?? url.match(GUID)?.[0];
  if (!id) return undefined;
  const appId = url.match(/apps\/([0-9a-f-]{36})/i)?.[1];
  const groupId = url.match(/groups\/([0-9a-f-]{36})/i)?.[1];
  return { id: id.toLowerCase(), appId: appId?.toLowerCase(), groupId: groupId && groupId !== "me" ? groupId.toLowerCase() : undefined };
}

/** Reads a report's details from the server (model, access) to keep it. */
export async function addReport(ref: { id: string; groupId?: string; appId?: string }) {
  const q = new URLSearchParams({ id: ref.id, ...(ref.groupId ? { groupId: ref.groupId } : {}), ...(ref.appId ? { appId: ref.appId } : {}) });
  const r = await get<ReportRef>(`/powerbi/report?${q}`);
  await save("report", { ...r, addedAt: Date.now() });
  return r;
}

function AddDialog({ kind, open, onOpenChange, onAdded }: { kind: "report" | "dataset"; open: boolean; onOpenChange: (o: boolean) => void; onAdded: (id: string) => void }) {
  const [catalog, setCatalog] = useState<CatalogItem[]>();
  const [error, setError] = useState<string>();
  const [query, setQuery] = useState("");
  const [busy, setBusy] = useState<string>();
  useEffect(() => {
    if (!open || catalog) return;
    get<CatalogItem[]>(kind === "report" ? "/powerbi/reports" : "/powerbi/datasets").then(setCatalog, (e) => setError(errorText(e)));
  }, [open, kind, catalog]);
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  const shown = (catalog ?? []).filter((c) => words.every((w) => `${c.name} ${c.workspace}`.toLowerCase().includes(w))).slice(0, 100);
  const choose = async (ref: { id: string; groupId?: string; appId?: string; name?: string }) => {
    setBusy(ref.id);
    try {
      if (kind === "report") await addReport(ref);
      else await save("dataset", { id: ref.id, groupId: ref.groupId, name: ref.name ?? "Semantic model", addedAt: Date.now() } as DatasetRef & { addedAt: number });
      onAdded(ref.id);
      onOpenChange(false);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  const link = query.includes("powerbi.com") || GUID.test(query) ? parseLink(query, kind) : undefined;
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-xl rounded-3xl">
        <DialogHeader>
          <DialogTitle>{kind === "report" ? "Add a report" : "Add a semantic model"}</DialogTitle>
        </DialogHeader>
        <div className="flex items-center gap-2 rounded-xl border px-3">
          <SearchIcon className="size-4 text-muted-foreground" />
          <input autoFocus value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search, or paste a Power BI link" className="h-10 flex-1 bg-transparent text-[14px] outline-none" />
        </div>
        {link && (
          <Button variant="dark" onClick={() => choose(link)} disabled={!!busy}>
            Add from the link
          </Button>
        )}
        <div className="max-h-[50vh] overflow-y-auto">
          {error && <div className="p-3 text-[13px] text-destructive">{error}</div>}
          {!catalog && !error && (
            <div className="flex items-center gap-2 p-3 text-[13px] text-muted-foreground">
              <Loader2Icon className="size-4 animate-spin" /> Loading your {kind === "report" ? "reports" : "models"}…
            </div>
          )}
          {shown.map((c) => (
            <button key={c.id} onClick={() => choose(c)} disabled={!!busy} className="flex w-full items-center gap-3 rounded-xl px-3 py-2 text-left hover:bg-muted">
              <div className="min-w-0 flex-1">
                <div className="truncate text-[14px] font-medium">{c.name}</div>
                <div className="truncate text-[12px] text-muted-foreground">{c.workspace}</div>
              </div>
              {busy === c.id && <Loader2Icon className="size-4 animate-spin" />}
            </button>
          ))}
          {catalog && !shown.length && <div className="p-3 text-[13px] text-muted-foreground">Nothing matches.</div>}
        </div>
      </DialogContent>
    </Dialog>
  );
}

/** The chosen report or model of a screen, with its list and "Add". */
export function usePicked<K extends "report" | "dataset">(kind: K, screen: string) {
  const items = useItems(kind);
  const [id, setId] = useLocal<string>(`picked:${screen}:${kind}`, "");
  const sorted = useMemo(() => [...(items ?? [])].sort((a, b) => a.name.localeCompare(b.name)), [items]);
  const current = sorted.find((x) => x.id === id) ?? (items && !id ? sorted[0] : undefined);
  return { items: sorted, current, select: setId, loading: !items };
}

export function ItemPicker<T extends { id: string; name: string }>({ kind, items, current, onSelect, label, extra }: { kind: "report" | "dataset"; items: T[]; current?: T; onSelect: (id: string) => void; label?: string; extra?: (item: T) => string | undefined }) {
  const [adding, setAdding] = useState(false);
  return (
    <>
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <button className="flex max-w-[420px] items-center gap-1.5 rounded-full px-2 py-1 text-[26px] font-light tracking-tight text-foreground hover:bg-muted">
            <span className="truncate">{current?.name ?? label ?? (kind === "report" ? "Choose a report" : "Choose a model")}</span>
            <ChevronDownIcon className="size-5 shrink-0 text-muted-foreground" />
          </button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="w-80">
          {items.map((it) => (
            <DropdownMenuItem key={it.id} onSelect={() => onSelect(it.id)} className="group">
              <CheckIcon className={cn("size-4", it.id === current?.id ? "opacity-100" : "opacity-0")} />
              <span className="min-w-0 flex-1 truncate">{it.name}</span>
              {extra?.(it) && <span className="text-[11px] text-muted-foreground">{extra(it)}</span>}
              <button
                className="opacity-0 group-hover:opacity-100"
                onClick={(e) => {
                  e.stopPropagation();
                  void remove(kind, it.id);
                }}
                aria-label="Remove from the list"
              >
                <Trash2Icon className="size-3.5" />
              </button>
            </DropdownMenuItem>
          ))}
          {items.length > 0 && <DropdownMenuSeparator />}
          <DropdownMenuItem onSelect={() => setAdding(true)}>
            <PlusIcon /> {kind === "report" ? "Add a report…" : "Add a semantic model…"}
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
      <AddDialog kind={kind} open={adding} onOpenChange={setAdding} onAdded={onSelect} />
    </>
  );
}

export { AddDialog };

/** "Add a report / model" as a button (empty states). */
export function AddButton({ kind, onAdded }: { kind: "report" | "dataset"; onAdded: (id: string) => void }) {
  const [open, setOpen] = useState(false);
  return (
    <>
      <Button variant="dark" size="sm" className="mt-3" onClick={() => setOpen(true)}>
        <PlusIcon /> {kind === "report" ? "Add a report" : "Add a semantic model"}
      </Button>
      <AddDialog kind={kind} open={open} onOpenChange={setOpen} onAdded={onAdded} />
    </>
  );
}
