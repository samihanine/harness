import { useEffect, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { ResourceTemplate } from "@modelcontextprotocol/server";
import { applyPatch } from "fast-json-patch";
import type { Operation } from "fast-json-patch";
import { ArrowDownIcon, ArrowUpIcon, BlocksIcon, CopyIcon, DownloadIcon, PlusIcon, SearchIcon, Trash2Icon, UploadCloudIcon, XIcon } from "lucide-react";
import { z } from "zod";
import { errorResult, textResult, useAgentFeatures, useLatest } from "@/agent/server";
import { downloadPbix, publishToMyWorkspace } from "@/builder/publish";
import { VisualPreview, useVisualData, visualPivot } from "@/builder/render";
import type { BuilderReport, BuilderVisual, Field, VisualType } from "@/builder/spec";
import { PALETTE, VISUALS, VISUAL_TYPES, hexId, newReport, newVisual, reportSchema, rolesOf, slot, visualSchema } from "@/builder/spec";
import { builderReports } from "@/builder/store";
import { NeedsDb } from "@/components/guards";
import { IconButton } from "@/components/icon-button";
import { EmptyState } from "@/components/page";
import { SearchSelect } from "@/components/search-select";
import { SimpleSelect } from "@/components/simple-select";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Input } from "@/components/ui/input";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Switch } from "@/components/ui/switch";
import { db, useDb } from "@/db/db";
import { useSources } from "@/lib/use-source";
import { cn } from "@/lib/utils";
import { embedVisual, resetEmbed } from "@/pbi/embed";
import { modelSummary } from "@/pbi/model";
import type { ReportContent } from "@/pbi/types";
import type { Source } from "@/query/engine";
import { runQuery, toCsv } from "@/query/engine";
import { AGGREGATES, parseField, runPivot } from "@/query/pivot";

export const Route = createFileRoute("/builder")({
  component: () => (
    <NeedsDb>
      <Builder />
    </NeedsDb>
  ),
});

const catalog = () =>
  VISUAL_TYPES.map((t) => `${t} (${VISUALS[t].label}): ${Object.entries(rolesOf(t)).map(([role, r]) => `${role}=${r.label} [${r.kind}${r.max ? `, max ${r.max}` : ""}]`).join(", ")}`).join("\n");

const zodMessage = (e: unknown) => (e instanceof z.ZodError ? z.prettifyError(e) : e);

function Builder() {
  const all = builderReports.use();
  const [currentId, setCurrentId] = useState<string | null>(() => localStorage.getItem("pbi-agent:builder"));
  const report = all.find((r) => r.id === currentId) ?? null;
  const { sources } = useSources(report ? [report.datasetId] : []);
  const source = sources[0];
  const [pageIndex, setPageIndex] = useState(0);
  const [creating, setCreating] = useState(false);
  const [real, setReal] = useState(false);
  const [task, setTask] = useState<{ step: string; error?: string } | null>(null);
  const latest = useLatest({ report, source, pageIndex });

  useEffect(() => {
    if (currentId) localStorage.setItem("pbi-agent:builder", currentId);
  }, [currentId]);

  const page = report?.pages[Math.min(pageIndex, (report?.pages.length ?? 1) - 1)];
  const setVisuals = (visuals: BuilderVisual[]) => report && page && builderReports.put({ ...report, pages: report.pages.map((p) => (p === page ? { ...p, visuals } : p)) });

  const publish = async () => {
    if (!report || !source) return;
    setTask({ step: "Publishing…" });
    try {
      const reportId = await publishToMyWorkspace(report, source.model, (step) => setTask({ step }));
      builderReports.put({ ...report, published: { reportId, at: new Date().toISOString() } });
      setReal(true);
      setTask(null);
    } catch (e) {
      setTask({ step: "", error: e instanceof Error ? e.message : String(e) });
    }
  };

  useAgentFeatures(
    report
      ? `Report builder: "${report.title}" on dataset ${source?.title ?? report.datasetId} (DAX). The report (builder://report) has pages of visuals; each visual has a type, title, fields by data role ({ref: "'Table'[Column]" | "[Measure]", aggregate?: sum|avg|count|distinctcount|min|max for columns in value roles}), filters ({ref, values}), format and position on a 1280×720 page. Visual types and roles: builder://visual-types. Example reports: builder://context/{reportId}. Verify each visual with check_visual.`
      : "Report builder: no report open. The user can create one (dataset + optional example reports).",
    (server) => {
      if (!report) return [];
      const current = () => builderReports.get(report.id)!;
      const change = (next: BuilderReport, summary: string) => {
        const before = current();
        builderReports.put(reportSchema.parse(next));
        return textResult(summary, { name: "patch_report", arguments: { operations: [{ op: "replace", path: "", value: before }] } });
      };
      const findVisual = (id: string) => current().pages.flatMap((p) => p.visuals).find((v) => v.id === id);
      return [
        server.registerResource("model", "builder://model", { title: "Dataset model", mimeType: "text/plain", annotations: { priority: 1 } }, async (uri) => ({
          contents: [{ uri: uri.href, text: latest.current.source ? modelSummary(latest.current.source.model) : "Loading…" }],
        })),
        server.registerResource("report", "builder://report", { title: "Report being built", mimeType: "application/json", annotations: { priority: 1 } }, async (uri) => ({
          contents: [{ uri: uri.href, text: JSON.stringify(current()) }],
        })),
        server.registerResource("types", "builder://visual-types", { title: "Visual types and data roles", mimeType: "text/plain", annotations: { priority: 1 } }, async (uri) => ({
          contents: [{ uri: uri.href, text: catalog() }],
        })),
        server.registerResource(
          "context",
          new ResourceTemplate("builder://context/{reportId}", {
            list: async () => ({ resources: current().contextReports.map((id) => ({ uri: `builder://context/${id}`, name: `Example report: ${String(db.row("reports", id)?.title ?? id)}` })) }),
          }),
          { title: "Example report (pages, visuals, fields, format)", mimeType: "application/json" },
          async (uri, { reportId }) => {
            const content = await db.readJson<ReportContent>(db.row("reports", String(reportId))?.content_json);
            const compact = content?.pages.map((p) => ({ page: p.displayName, visuals: p.visuals.map(({ raw: _, ...v }) => v) }));
            return { contents: [{ uri: uri.href, text: JSON.stringify(compact ?? null) }] };
          },
        ),
        server.registerTool(
          "add_visual",
          {
            description: "Adds a visual (one row in the builder). Position defaults to the next free slot.",
            inputSchema: z.object({ page: z.number().int().optional().describe("page index (default: current)"), visual: visualSchema.partial({ id: true, x: true, y: true, width: true, height: true }) }),
          },
          async ({ page: index, visual }) => {
            const r = current();
            const target = index ?? Math.min(latest.current.pageIndex, r.pages.length - 1);
            if (!r.pages[target]) return errorResult(`No page ${target}`);
            const id = visual.id ?? hexId();
            try {
              const created = visualSchema.parse({ ...slot(r.pages[target].visuals.length), ...visual, id });
              return change({ ...r, pages: r.pages.map((p, i) => (i === target ? { ...p, visuals: [...p.visuals, created] } : p)) }, `Visual ${id} added.`);
            } catch (e) {
              return errorResult(zodMessage(e));
            }
          },
        ),
        server.registerTool(
          "update_visual",
          { description: "Changes a visual: given keys replace the current ones (fields and format are merged by key).", inputSchema: z.object({ id: z.string(), patch: visualSchema.partial().omit({ id: true }) }) },
          async ({ id, patch }) => {
            const r = current();
            if (!findVisual(id)) return errorResult(`No visual ${id}`);
            try {
              const update = (v: BuilderVisual) => visualSchema.parse({ ...v, ...patch, fields: { ...v.fields, ...patch.fields }, format: { ...v.format, ...patch.format } });
              return change({ ...r, pages: r.pages.map((p) => ({ ...p, visuals: p.visuals.map((v) => (v.id === id ? update(v) : v)) })) }, "Visual updated.");
            } catch (e) {
              return errorResult(zodMessage(e));
            }
          },
        ),
        server.registerTool("remove_visual", { description: "Removes a visual.", inputSchema: z.object({ id: z.string() }) }, async ({ id }) => {
          const r = current();
          if (!findVisual(id)) return errorResult(`No visual ${id}`);
          return change({ ...r, pages: r.pages.map((p) => ({ ...p, visuals: p.visuals.filter((v) => v.id !== id) })) }, "Visual removed.");
        }),
        server.registerTool(
          "patch_report",
          {
            description: "Any other change (pages, order, title…) with JSON Patch operations on builder://report.",
            inputSchema: z.object({ operations: z.array(z.object({ op: z.string(), path: z.string(), value: z.unknown().optional(), from: z.string().optional() })) }),
          },
          async ({ operations }) => {
            try {
              const next = applyPatch(structuredClone(current()), operations as Operation[], true, false).newDocument;
              return change({ ...next, id: report.id }, "Report updated.");
            } catch (e) {
              return errorResult(zodMessage(e));
            }
          },
        ),
        server.registerTool(
          "check_visual",
          { description: "Runs the query of a visual and returns its data (or the error), to verify it.", inputSchema: z.object({ id: z.string() }), annotations: { readOnlyHint: true } },
          async ({ id }) => {
            const visual = findVisual(id);
            const src = latest.current.source;
            if (!visual || !src) return errorResult(`No visual ${id}`);
            const { pivot } = visualPivot(visual, src);
            if (!pivot.values.length) return textResult("No value field: nothing to compute.");
            const result = await runPivot(pivot, [src]);
            const error = result.queries.find((q) => q.error);
            return error ? errorResult(`${error.error}\nQuery:\n${error.query}`) : textResult(`${result.rows.length} rows\n${JSON.stringify(result.rows.slice(0, 30))}`);
          },
        ),
        server.registerTool(
          "run_query",
          { description: "Runs a DAX query on the dataset (exploration); returns CSV.", inputSchema: z.object({ query: z.string() }), annotations: { readOnlyHint: true } },
          async ({ query }) => {
            try {
              const r = await runQuery(latest.current.source!, query);
              return textResult(`${r.rows.length} rows\n${toCsv(r, 100)}`);
            } catch (e) {
              return errorResult(e);
            }
          },
        ),
      ];
    },
    [report?.id],
  );

  return (
    <div className="flex h-full flex-col">
      <div className="flex h-11 shrink-0 items-center gap-1.5 border-b px-3">
        <SearchSelect size="sm" className="w-56" value={report?.id ?? null} placeholder="Open a report…" onChange={setCurrentId} onAdd={() => setCreating(true)} addLabel="New report" options={all.map((r) => ({ value: r.id, label: r.title }))} />
        <IconButton label="New report" onClick={() => setCreating(true)}>
          <PlusIcon />
        </IconButton>
        {report && (
          <>
            <div className="ml-2 flex items-center gap-0.5">
              {report.pages.map((p, i) => (
                <button
                  key={p.name}
                  type="button"
                  onClick={() => setPageIndex(i)}
                  className={cn("rounded-md px-2 py-1 text-[12px] text-muted-foreground hover:text-foreground", i === pageIndex && "bg-muted text-foreground")}
                >
                  {p.displayName}
                </button>
              ))}
              <IconButton
                label="Add page"
                size="icon-xs"
                onClick={() => {
                  builderReports.put({ ...report, pages: [...report.pages, { name: `ReportSection${hexId()}`, displayName: `Page ${report.pages.length + 1}`, visuals: [] }] });
                  setPageIndex(report.pages.length);
                }}
              >
                <PlusIcon />
              </IconButton>
            </div>
            <span className="ml-auto" />
            {task && (
              <span className={cn("max-w-xs truncate text-[12px]", task.error ? "text-destructive" : "shimmer")} title={task.error}>
                {task.error ?? task.step}
              </span>
            )}
            {report.published && (
              <label className="flex items-center gap-1.5 text-[12px] text-muted-foreground">
                <Switch size="sm" checked={real} onCheckedChange={setReal} /> Power BI
              </label>
            )}
            <IconButton label={report.published ? "Update the Power BI copy (My workspace)" : "Open in Power BI (imports a copy in My workspace)"} disabled={!source} onClick={() => void publish()}>
              <UploadCloudIcon />
            </IconButton>
            <IconButton label="Download .pbix (live connection)" disabled={!source} onClick={() => source && downloadPbix(report, source.model)}>
              <DownloadIcon />
            </IconButton>
            <DropdownMenu>
              <DropdownMenuTrigger render={<Button size="sm" variant="outline" />}>
                <PlusIcon /> Visual
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="w-52">
                {VISUAL_TYPES.map((t) => (
                  <DropdownMenuItem key={t} onClick={() => page && setVisuals([...page.visuals, newVisual(t, page.visuals.length)])}>
                    {VISUALS[t].label}
                  </DropdownMenuItem>
                ))}
              </DropdownMenuContent>
            </DropdownMenu>
          </>
        )}
      </div>

      {!report ? (
        <div className="mx-auto w-full max-w-lg p-8">
          <EmptyState icon={<BlocksIcon />} title="Build a report">
            <p className="mb-3">Pick a dataset (and example reports if you want), then add visuals yourself or with the agent.</p>
            <Button size="sm" onClick={() => setCreating(true)}>
              New report
            </Button>
          </EmptyState>
        </div>
      ) : (
        <div className="min-h-0 flex-1 overflow-y-auto">
          {page?.visuals.length === 0 && <p className="p-8 text-center text-[12px] text-muted-foreground">No visual on this page yet: add one, or ask the agent.</p>}
          {page?.visuals.map((visual, i) => (
            <VisualRow
              key={visual.id}
              visual={visual}
              source={source}
              published={real && report.published ? { reportId: report.published.reportId, pageName: page.name } : undefined}
              onChange={(v) => setVisuals(page.visuals.map((x) => (x.id === v.id ? v : x)))}
              onMove={(delta) => {
                const next = [...page.visuals];
                next.splice(i + delta, 0, ...next.splice(i, 1));
                setVisuals(next);
              }}
              onDuplicate={() => setVisuals([...page.visuals, { ...structuredClone(visual), id: hexId(), ...slot(page.visuals.length) }])}
              onRemove={() => setVisuals(page.visuals.filter((x) => x.id !== visual.id))}
              first={i === 0}
              last={i === page.visuals.length - 1}
            />
          ))}
        </div>
      )}
      <NewReportDialog
        open={creating}
        onClose={() => setCreating(false)}
        onCreate={(r) => {
          builderReports.put(r);
          setCurrentId(r.id);
          setPageIndex(0);
        }}
      />
    </div>
  );
}

function VisualRow({
  visual,
  source,
  published,
  onChange,
  onMove,
  onDuplicate,
  onRemove,
  first,
  last,
}: {
  visual: BuilderVisual;
  source?: Source;
  published?: { reportId: string; pageName: string };
  onChange: (visual: BuilderVisual) => void;
  onMove: (delta: number) => void;
  onDuplicate: () => void;
  onRemove: () => void;
  first: boolean;
  last: boolean;
}) {
  return (
    <div className="flex border-b animate-in fade-in-0">
      <div className="min-w-0 flex-1 p-4">
        <div className="h-80 rounded-xl border bg-card shadow-soft">
          {published ? <RealVisual {...published} visualName={visual.id} /> : <VisualPreview visual={visual} source={source} />}
        </div>
      </div>
      <div className="w-80 shrink-0 overflow-y-auto border-l bg-muted/20" style={{ maxHeight: 352 }}>
        <div className="sticky top-0 z-10 flex items-center gap-1 border-b bg-background/90 px-2 py-1.5 backdrop-blur">
          <SimpleSelect size="sm" className="flex-1" value={visual.type} onChange={(type) => onChange({ ...visual, type: type as VisualType })} options={VISUAL_TYPES.map((t) => ({ value: t, label: VISUALS[t].label }))} />
          <IconButton label="Move up" size="icon-xs" disabled={first} onClick={() => onMove(-1)}>
            <ArrowUpIcon />
          </IconButton>
          <IconButton label="Move down" size="icon-xs" disabled={last} onClick={() => onMove(1)}>
            <ArrowDownIcon />
          </IconButton>
          <IconButton label="Duplicate" size="icon-xs" onClick={onDuplicate}>
            <CopyIcon />
          </IconButton>
          <IconButton label="Delete" size="icon-xs" onClick={onRemove}>
            <Trash2Icon />
          </IconButton>
        </div>
        <Properties visual={visual} source={source} onChange={onChange} />
      </div>
    </div>
  );
}

function RealVisual({ reportId, pageName, visualName }: { reportId: string; pageName: string; visualName: string }) {
  const element = useRef<HTMLDivElement>(null);
  const [error, setError] = useState("");
  useEffect(() => {
    const el = element.current;
    if (!el) return;
    embedVisual(el, { reportId }, pageName, visualName).catch((e: Error) => setError(e.message));
    return () => resetEmbed(el);
  }, [reportId, pageName, visualName]);
  return error ? <p className="p-3 text-[12px] text-destructive">{error} (update the Power BI copy)</p> : <div ref={element} className="size-full" />;
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-1.5 border-b px-3 py-2.5">
      <span className="text-[10px] font-medium tracking-wide text-muted-foreground uppercase">{title}</span>
      {children}
    </div>
  );
}

function Properties({ visual, source, onChange }: { visual: BuilderVisual; source?: Source; onChange: (visual: BuilderVisual) => void }) {
  const roles = rolesOf(visual.type);
  const fmt = visual.format;
  const setFormat = (patch: Partial<BuilderVisual["format"]>) => onChange({ ...visual, format: { ...fmt, ...patch } });
  const setFields = (role: string, fields: Field[]) => onChange({ ...visual, fields: { ...visual.fields, [role]: fields } });
  const { error } = useVisualData(visual, source);
  return (
    <>
      <Section title="Title">
        <Input value={visual.title} onChange={(e) => onChange({ ...visual, title: e.target.value })} className="h-7" />
      </Section>
      <Section title="Data">
        {Object.entries(roles).map(([role, info]) => {
          const fields = visual.fields[role] ?? [];
          return (
            <div key={role} className="flex flex-col gap-1">
              <span className="text-[11px] text-muted-foreground">{info.label}</span>
              {fields.map((f, i) => (
                <div key={i} className="flex h-6 items-center gap-1 rounded-md border bg-card pl-2 text-[12px]">
                  <span className="min-w-0 flex-1 truncate" title={f.ref}>
                    {parseField(f.ref).column}
                  </span>
                  {parseField(f.ref).table && info.kind !== "dimension" && (
                    <select
                      value={f.aggregate ?? (info.kind === "value" ? "sum" : "")}
                      onChange={(e) => setFields(role, fields.map((x, j) => (j === i ? { ...x, aggregate: (e.target.value || undefined) as Field["aggregate"] } : x)))}
                      className="bg-transparent text-[11px] text-muted-foreground outline-none"
                    >
                      {info.kind === "any" && <option value="">none</option>}
                      {AGGREGATES.map((a) => (
                        <option key={a}>{a}</option>
                      ))}
                    </select>
                  )}
                  <button type="button" aria-label="Remove" onClick={() => setFields(role, fields.filter((_, j) => j !== i))} className="px-1 text-muted-foreground hover:text-foreground">
                    <XIcon className="size-3" />
                  </button>
                </div>
              ))}
              {(!info.max || fields.length < info.max) && source && <FieldPicker source={source} kind={info.kind} onPick={(f) => setFields(role, [...fields, f])} />}
            </div>
          );
        })}
        {error && <p className="text-[11px] text-destructive">{error}</p>}
      </Section>
      <Section title="Filters">
        {visual.filters.map((f, i) => (
          <div key={i} className="flex flex-col gap-1">
            <span className="flex items-center text-[11px] text-muted-foreground">
              {parseField(f.ref).column}
              <button type="button" aria-label="Remove" className="ml-auto" onClick={() => onChange({ ...visual, filters: visual.filters.filter((_, j) => j !== i) })}>
                <XIcon className="size-3" />
              </button>
            </span>
            <Input
              defaultValue={f.values.join("; ")}
              placeholder="Values separated by ;"
              onBlur={(e) =>
                onChange({
                  ...visual,
                  filters: visual.filters.map((x, j) => (j === i ? { ...x, values: e.target.value.split(";").map((v) => v.trim()).filter(Boolean) } : x)),
                })
              }
              className="h-7"
            />
          </div>
        ))}
        {source && <FieldPicker source={source} kind="dimension" label="Add filter" onPick={(f) => onChange({ ...visual, filters: [...visual.filters, { ref: f.ref, values: [] }] })} />}
      </Section>
      <Section title="Format">
        <Toggle label="Title" checked={fmt.showTitle} onChange={(showTitle) => setFormat({ showTitle })} />
        <Toggle label="Legend" checked={fmt.showLegend} onChange={(showLegend) => setFormat({ showLegend })} />
        {fmt.showLegend && (
          <SimpleSelect size="sm" value={fmt.legendPosition} onChange={(p) => setFormat({ legendPosition: p as "Top" })} options={["Top", "Bottom", "Left", "Right"].map((p) => ({ value: p, label: `Legend ${p.toLowerCase()}` }))} />
        )}
        <Toggle label="Data labels" checked={fmt.showDataLabels} onChange={(showDataLabels) => setFormat({ showDataLabels })} />
        <Toggle label="Axis titles" checked={fmt.showAxisTitles} onChange={(showAxisTitles) => setFormat({ showAxisTitles })} />
        <div className="flex items-center gap-2 text-[12px]">
          <span className="flex-1">Colors</span>
          {[0, 1, 2].map((i) => (
            <input
              key={i}
              type="color"
              value={fmt.colors[i] ?? PALETTE[i]}
              onChange={(e) => {
                const colors = [...fmt.colors];
                for (let j = colors.length; j < i; j++) colors[j] = PALETTE[j];
                colors[i] = e.target.value;
                setFormat({ colors });
              }}
              className="size-5 cursor-pointer rounded border-0 bg-transparent p-0"
            />
          ))}
        </div>
        <div className="flex items-center gap-2 text-[12px]">
          <span className="flex-1">Background</span>
          <input type="color" value={fmt.background ?? "#ffffff"} onChange={(e) => setFormat({ background: e.target.value })} className="size-5 cursor-pointer rounded border-0 bg-transparent p-0" />
          {fmt.background && (
            <button type="button" aria-label="No background" onClick={() => setFormat({ background: undefined })}>
              <XIcon className="size-3 text-muted-foreground" />
            </button>
          )}
        </div>
        <div className="grid grid-cols-2 gap-2">
          <NumberInput label="Title size" value={fmt.titleSize} onChange={(titleSize) => setFormat({ titleSize })} />
          <NumberInput label="Font size" value={fmt.fontSize} onChange={(fontSize) => setFormat({ fontSize })} />
        </div>
      </Section>
      <Section title="Position (1280 × 720)">
        <div className="grid grid-cols-4 gap-1.5">
          {(["x", "y", "width", "height"] as const).map((k) => (
            <NumberInput key={k} label={k} value={visual[k]} onChange={(v) => onChange({ ...visual, [k]: v ?? 0 })} />
          ))}
        </div>
      </Section>
    </>
  );
}

function Toggle({ label, checked, onChange }: { label: string; checked: boolean; onChange: (checked: boolean) => void }) {
  return (
    <label className="flex items-center justify-between text-[12px]">
      {label}
      <Switch size="sm" checked={checked} onCheckedChange={onChange} />
    </label>
  );
}

function NumberInput({ label, value, onChange }: { label: string; value?: number; onChange: (value: number | undefined) => void }) {
  return (
    <label className="flex flex-col gap-0.5 text-[10px] text-muted-foreground">
      {label}
      <input
        type="number"
        value={value ?? ""}
        onChange={(e) => onChange(e.target.value === "" ? undefined : Number(e.target.value))}
        className="h-6 rounded-md border bg-background px-1.5 text-[12px] text-foreground outline-none focus:border-ring"
      />
    </label>
  );
}

/** Searchable list of the model fields. */
function FieldPicker({ source, kind, label = "Add field", onPick }: { source: Source; kind: "dimension" | "value" | "any"; label?: string; onPick: (field: Field) => void }) {
  const [search, setSearch] = useState("");
  const q = search.toLowerCase();
  const fields = source.model.tables.flatMap((t) => [
    ...(kind === "dimension" ? [] : t.measures.map((m) => ({ ref: `[${m.name}]`, label: m.name, table: t.name, measure: true }))),
    ...t.columns.filter((c) => !c.hidden).map((c) => ({ ref: `'${t.name}'[${c.name}]`, label: c.name, table: t.name, measure: false })),
  ]);
  const shown = fields.filter((f) => !q || f.label.toLowerCase().includes(q) || f.table.toLowerCase().includes(q)).slice(0, 80);
  return (
    <Popover onOpenChange={() => setSearch("")}>
      <PopoverTrigger className="flex h-6 items-center gap-1 rounded-md border border-dashed px-2 text-[12px] text-muted-foreground hover:text-foreground">
        <PlusIcon className="size-3" /> {label}
      </PopoverTrigger>
      <PopoverContent align="start" className="w-72 gap-0 p-1">
        <label className="flex items-center gap-1.5 border-b px-2 pb-1">
          <SearchIcon className="size-3.5 text-muted-foreground" />
          <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search" className="h-7 flex-1 bg-transparent outline-none" />
        </label>
        <div className="max-h-72 overflow-y-auto pt-1">
          {shown.map((f) => (
            <button
              key={f.ref}
              type="button"
              onClick={() => onPick({ ref: f.ref, aggregate: !f.measure && kind === "value" ? "sum" : undefined })}
              className="flex w-full items-center gap-2 rounded-md px-2 py-1 text-left text-[12px] hover:bg-muted"
            >
              <span className={cn("size-1.5 shrink-0 rounded-full", f.measure ? "bg-amber-500" : "bg-zinc-400")} />
              <span className="truncate">{f.label}</span>
              <span className="ml-auto truncate text-[10px] text-muted-foreground">{f.table}</span>
            </button>
          ))}
        </div>
      </PopoverContent>
    </Popover>
  );
}

function NewReportDialog({ open, onClose, onCreate }: { open: boolean; onClose: () => void; onCreate: (report: BuilderReport) => void }) {
  const data = useDb();
  const [title, setTitle] = useState("");
  const [datasetId, setDatasetId] = useState<string | null>(null);
  const [examples, setExamples] = useState<string[]>([]);
  const remote = data.rows("datasets").filter((d) => d.source !== "excel");
  return (
    <Dialog open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="gap-3 sm:max-w-md">
        <DialogTitle className="text-[13px] font-medium">New report</DialogTitle>
        <Input autoFocus value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Title" />
        <SearchSelect value={datasetId} onChange={setDatasetId} placeholder="Dataset (published semantic model)" options={remote.map((d) => ({ value: d.id, label: String(d.title), recent: String(d.updated_at ?? "") }))} />
        {remote.length === 0 && <p className="text-[11px] text-muted-foreground">The builder needs a published dataset: add a report using it in the database.</p>}
        <div className="flex flex-col gap-1">
          <span className="text-[12px] font-medium">Example reports (optional, context for the agent)</span>
          <div className="max-h-40 overflow-y-auto rounded-lg border p-1">
            {data.rows("reports").map((r) => (
              <label key={r.id} className="flex items-center gap-2 rounded-md px-2 py-1 hover:bg-muted">
                <Checkbox checked={examples.includes(r.id)} onCheckedChange={(on) => setExamples(on ? [...examples, r.id] : examples.filter((id) => id !== r.id))} />
                <span className="truncate">{String(r.title)}</span>
              </label>
            ))}
            {data.rows("reports").length === 0 && <p className="p-2 text-[11px] text-muted-foreground">No report in the database.</p>}
          </div>
        </div>
        <div className="flex justify-end gap-2">
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            disabled={!title.trim() || !datasetId}
            onClick={() => {
              onCreate({ ...newReport(title.trim(), datasetId!), contextReports: examples });
              setTitle("");
              setExamples([]);
              onClose();
            }}
          >
            Create
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}
