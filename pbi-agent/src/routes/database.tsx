import { useEffect, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { ResourceTemplate } from "@modelcontextprotocol/server";
import { BracesIcon, DatabaseIcon, FileUpIcon, LinkIcon, PlusIcon, RefreshCwIcon, SheetIcon, Trash2Icon } from "lucide-react";
import { z } from "zod";
import { errorResult, textResult, useAgentFeatures } from "@/agent/server";
import { NeedsDb } from "@/components/guards";
import { IconButton } from "@/components/icon-button";
import { EmptyState, ago } from "@/components/page";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Input } from "@/components/ui/input";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { db, useDb } from "@/db/db";
import type { TableName } from "@/db/db";
import { addExcelDataset, addReportFromPbix, addReportFromUrl, refreshDataset, refreshReport } from "@/db/ops";
import type { Row } from "@/db/schema";
import { TINTS } from "@/lib/colors";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/database")({
  component: () => (
    <NeedsDb>
      <Database />
    </NeedsDb>
  ),
  validateSearch: (search: Record<string, unknown>): { tab?: TableName } =>
    ["reports", "datasets", "guides"].includes(String(search.tab)) ? { tab: search.tab as TableName } : {},
});

/** Light columns of a row for lists (heavy JSON stays in files). */
const light = (table: TableName, row: Row) =>
  table === "guides"
    ? { id: row.id, title: row.title, visual_ids: row.visual_ids, content: String(row.content ?? "").slice(0, 300) }
    : { id: row.id, title: row.title, ...(table === "reports" ? { dataset_id: row.dataset_id } : { source: row.source }), context: row.context, updated_at: row.updated_at };

function Database() {
  const tab = Route.useSearch().tab ?? "reports";
  const navigate = Route.useNavigate();
  const data = useDb();
  const [task, setTask] = useState<{ step: string; error?: string } | null>(null);
  const [linkOpen, setLinkOpen] = useState(false);
  const [json, setJson] = useState<{ title: string; value: unknown } | null>(null);
  const pbix = useRef<HTMLInputElement>(null);

  const run = async (action: (log: (step: string) => void) => Promise<unknown>) => {
    setTask({ step: "Working…" });
    try {
      await action((step) => setTask({ step }));
      setTask(null);
    } catch (error) {
      setTask({ step: "", error: error instanceof Error ? error.message : String(error) });
    }
  };

  useAgentFeatures(
    "Database page: lists of reports, datasets (semantic models) and guides stored in local Excel files. Each content_json (report pages/visuals, dataset model) is readable with the db://{table}/{id}/content template. You can rename rows and edit their context (a free text helping you later).",
    (server) => [
      ...(["reports", "datasets", "guides"] as TableName[]).map((table) =>
        server.registerResource(
          table,
          `db://${table}`,
          { title: `${table} (one row per line)`, mimeType: "application/jsonl", annotations: { priority: table === "guides" ? 0.5 : 1 } },
          async (uri) => ({ contents: [{ uri: uri.href, text: db.rows(table).map((r) => JSON.stringify(light(table, r))).join("\n") }] }),
        ),
      ),
      server.registerResource(
        "content",
        new ResourceTemplate("db://{table}/{id}/content", { list: undefined }),
        { title: "content_json of a report or dataset", mimeType: "application/json" },
        async (uri, { table, id }) => {
          const row = db.row(table as TableName, String(id));
          const content = row ? await db.readJson(row.content_json) : null;
          return { contents: [{ uri: uri.href, text: JSON.stringify(content ?? null, null, 1) }] };
        },
      ),
      server.registerTool(
        "update_database_row",
        {
          description: "Renames a report / dataset / guide or changes its context.",
          inputSchema: z.object({ table: z.enum(["reports", "datasets", "guides"]), id: z.string(), title: z.string().optional(), context: z.string().optional() }),
        },
        async ({ table, id, ...values }) => {
          const row = db.row(table, id);
          if (!row) return errorResult(`No ${table} row ${id}`);
          const patch = Object.fromEntries(Object.entries(values).filter(([k, v]) => v !== undefined && (table !== "guides" || k === "title")));
          const before = Object.fromEntries(Object.keys(patch).map((k) => [k, row[k] ?? ""]));
          db.tables[table].update([{ id, values: patch }]);
          return textResult("Updated.", { name: "update_database_row", arguments: { table, id, ...before } });
        },
      ),
    ],
    [],
  );

  const rows = data.rows(tab);
  return (
    <div className="mx-auto flex max-w-5xl flex-col gap-4 px-6 py-6 animate-in fade-in-0">
      <div className="flex items-center gap-2">
        <div className="flex rounded-lg bg-muted p-0.5">
          {(["reports", "datasets", "guides"] as TableName[]).map((t) => (
            <button
              key={t}
              type="button"
              onClick={() => void navigate({ search: { tab: t } })}
              className={cn("rounded-md px-2.5 py-1 text-muted-foreground capitalize transition-all", tab === t && "bg-background text-foreground shadow-soft")}
            >
              {t} <span className="ml-0.5 text-[11px] tabular-nums opacity-60">{data.rows(t).length}</span>
            </button>
          ))}
        </div>
        <span className="ml-auto" />
        {task && (
          <span className={cn("max-w-md truncate text-[12px]", task.error ? "text-destructive" : "shimmer")} title={task.error}>
            {task.error ?? task.step}
          </span>
        )}
        {tab === "reports" && (
          <DropdownMenu>
            <DropdownMenuTrigger render={<Button size="sm" variant="outline" />}>
              <PlusIcon /> Add report
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" className="w-56">
              <DropdownMenuItem onClick={() => pbix.current?.click()}>
                <FileUpIcon /> From a .pbix file…
              </DropdownMenuItem>
              <DropdownMenuItem onClick={() => setLinkOpen(true)}>
                <LinkIcon /> From a link…
              </DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
        )}
        {tab === "datasets" && (
          <Button
            size="sm"
            variant="outline"
            onClick={() =>
              void window
                .showOpenFilePicker({ types: [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }] })
                .then(([handle]) => run((log) => addExcelDataset(handle, log)))
                .catch(() => undefined)
            }
          >
            <SheetIcon /> Add from Excel
          </Button>
        )}
        <input ref={pbix} type="file" accept=".pbix" hidden onChange={(e) => e.target.files?.[0] && void run((log) => addReportFromPbix(e.target.files![0], log))} />
      </div>

      {rows.length === 0 ? (
        <EmptyState icon={<DatabaseIcon />} title={`No ${tab} yet`}>
          {tab === "reports" && "Add a report from its .pbix file (full definition) or its link: its dataset is added automatically."}
          {tab === "datasets" && "Datasets come with the reports using them, or from an Excel file (local model)."}
          {tab === "guides" && "Guides are written from the Viewer, next to the visuals they explain."}
        </EmptyState>
      ) : (
        <div className="overflow-hidden rounded-xl border bg-card shadow-soft">
          <table className="w-full">
            <thead>
              <tr className="border-b text-left text-[12px] text-muted-foreground">
                <th className="px-3 py-2 font-medium">Title</th>
                <th className="px-3 py-2 font-medium">{tab === "reports" ? "Dataset" : tab === "datasets" ? "Source" : "Visuals"}</th>
                <th className="px-3 py-2 font-medium">{tab === "guides" ? "Content" : "Context"}</th>
                <th className="px-3 py-2 font-medium">{tab === "guides" ? "" : "Updated"}</th>
                <th className="w-24" />
              </tr>
            </thead>
            <tbody className="divide-y">
              {rows.map((row) => (
                <tr key={row.id} className="group align-top hover:bg-muted/30">
                  <td className="px-1 py-1">
                    <EditableText value={String(row.title ?? "")} onSave={(title) => db.tables[tab].update([{ id: row.id, values: { title } }])} />
                    <span className="block px-2 font-mono text-[10px] text-muted-foreground/60">{row.id}</span>
                  </td>
                  <td className="px-3 py-2 text-[12px] text-muted-foreground">
                    {tab === "reports" && String(data.row("datasets", String(row.dataset_id))?.title ?? (row.dataset_id ? String(row.dataset_id).slice(0, 8) : "—"))}
                    {tab === "datasets" && <span className={cn("rounded-full px-2 py-0.5", TINTS[row.source === "excel" ? "green" : "blue"])}>{String(row.source ?? "remote")}</span>}
                    {tab === "guides" && `${String(row.visual_ids ?? "").split(/[;\n]/).filter(Boolean).length} linked`}
                  </td>
                  <td className="max-w-xs px-1 py-1">
                    {tab === "guides" ? (
                      <p className="line-clamp-2 px-2 py-1 text-[12px] text-muted-foreground">{String(row.content ?? "")}</p>
                    ) : (
                      <ContextEditor value={String(row.context ?? "")} onSave={(context) => db.tables[tab].update([{ id: row.id, values: { context } }])} />
                    )}
                  </td>
                  <td className="px-3 py-2 text-[12px] whitespace-nowrap text-muted-foreground">{row.updated_at ? ago(Date.parse(String(row.updated_at))) : ""}</td>
                  <td className="px-2 py-1">
                    <div className="flex justify-end opacity-0 transition-opacity group-hover:opacity-100">
                      {tab !== "guides" && (
                        <>
                          <IconButton label="View content_json" onClick={() => void db.readJson(row.content_json).then((value) => setJson({ title: String(row.title), value }))}>
                            <BracesIcon />
                          </IconButton>
                          <IconButton
                            label="Refresh content"
                            onClick={() =>
                              void run((log) =>
                                tab === "reports" ? refreshReport({ reportId: row.id, groupId: (row.workspace_id as string) || undefined }, log) : refreshDataset(row.id, log),
                              )
                            }
                          >
                            <RefreshCwIcon />
                          </IconButton>
                        </>
                      )}
                      <ConfirmDelete onConfirm={() => void db.remove(tab, row.id)} />
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <LinkDialog open={linkOpen} onClose={() => setLinkOpen(false)} onSubmit={(url) => void run((log) => addReportFromUrl(url, log))} />
      <Dialog open={!!json} onOpenChange={(o) => !o && setJson(null)}>
        <DialogContent className="flex h-[80dvh] flex-col gap-2 sm:max-w-3xl">
          <DialogTitle className="text-[13px] font-medium">{json?.title}</DialogTitle>
          <pre className="min-h-0 flex-1 overflow-auto rounded-lg border bg-muted/40 p-3 font-mono text-[11px]">{JSON.stringify(json?.value ?? null, null, 2)}</pre>
        </DialogContent>
      </Dialog>
    </div>
  );
}

function EditableText({ value, onSave }: { value: string; onSave: (value: string) => void }) {
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);
  return (
    <input
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => draft.trim() && draft !== value && onSave(draft.trim())}
      onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
      className="h-7 w-full rounded-md bg-transparent px-2 font-medium outline-none focus:bg-background focus:ring-1 focus:ring-ring/60"
    />
  );
}

function ContextEditor({ value, onSave }: { value: string; onSave: (value: string) => void }) {
  const [draft, setDraft] = useState(value);
  return (
    <Popover onOpenChange={(open) => (open ? setDraft(value) : draft !== value && onSave(draft))}>
      <PopoverTrigger className="w-full rounded-md px-2 py-1 text-left text-[12px] hover:bg-muted">
        {value ? <span className="line-clamp-2">{value}</span> : <span className="text-muted-foreground/60">Add context…</span>}
      </PopoverTrigger>
      <PopoverContent align="start" className="w-96 p-2">
        <textarea
          autoFocus
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          placeholder="What it is for, caveats, how to read it… (given to the agent)"
          className="field-sizing-content min-h-24 w-full resize-none rounded-md border bg-transparent p-2 text-[12px] outline-none focus:border-ring"
        />
      </PopoverContent>
    </Popover>
  );
}

function ConfirmDelete({ onConfirm }: { onConfirm: () => void }) {
  return (
    <Popover>
      <PopoverTrigger render={<IconButton label="Delete" />}>
        <Trash2Icon />
      </PopoverTrigger>
      <PopoverContent align="end" className="w-56 gap-2 p-2.5">
        <p className="text-[12px]">Delete this row and its stored JSON?</p>
        <Button size="sm" variant="destructive" onClick={onConfirm}>
          Delete
        </Button>
      </PopoverContent>
    </Popover>
  );
}

function LinkDialog({ open, onClose, onSubmit }: { open: boolean; onClose: () => void; onSubmit: (url: string) => void }) {
  const [url, setUrl] = useState("");
  return (
    <Dialog open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="gap-3 sm:max-w-lg">
        <DialogTitle className="text-[13px] font-medium">Add a report from its link</DialogTitle>
        <Input autoFocus value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://app.powerbi.com/groups/…/reports/…" />
        <p className="text-[11px] text-muted-foreground">
          The full definition is read when you may download the report; otherwise it is read through embedding (formatting partial).
        </p>
        <div className="flex justify-end gap-2">
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            disabled={!url.trim()}
            onClick={() => {
              onSubmit(url.trim());
              setUrl("");
              onClose();
            }}
          >
            Add
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}
