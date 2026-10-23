import { useEffect, useMemo, useRef, useState } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import { BookOpenIcon, BookTextIcon, DatabaseIcon, ChartPieIcon, CheckIcon, CopyIcon, DownloadIcon, PencilIcon, PlusIcon, SparklesIcon } from "lucide-react";
import { withToast } from "@/components/entity";
import { Markdown } from "@/components/markdown";
import { ReportPicker } from "@/components/pickers";
import { ReportActions } from "@/components/report-actions";
import { RichText } from "@/components/rich-text";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { ContextLinks } from "@/components/context-links";
import { DataTable } from "@/components/data-table";
import { Empty, Section, Workspace } from "@/components/workspace";
import { Zoom } from "@/components/zoom";
import { EXPORT, GUIDE_EDITOR, VIEWER_TABS, VISUAL_EXPLAIN, VISUAL_PREVIEW } from "@/config";
import { reportContext } from "@/agent/context";
import { explainVisual } from "@/agent/explain";
import { describeError, type Agent } from "@/agent/loop";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES, bookmarksText, daxTool, viewerTools } from "@/agent/pbi-tools";
import { useLast } from "@/lib/last";
import { embedVisual, models, pageText, setVisualFilters, type Ref, type Report } from "@/lib/pbi";
import type { Row } from "@/lib/schema";
import { addRows, updateRows, withRows, withUpdates } from "@/lib/excel";
import { find, patch, useCollection, type DatasetEntry, type ReportEntry, type SchemaEntry } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";
import { useTable } from "@/lib/use-table";
import { cn } from "@/lib/utils";
import { capturePreviewPng, captureVisualPng, printVisual } from "@/lib/visual-image";

export const Route = createFileRoute("/viewer")({
  validateSearch: (s: Record<string, unknown>): { report?: string; tab?: string } => ({ report: (s.report as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Viewer,
});

/** Rows of the built-in Guides schema for one report: none until its Excel file is set. */
function useReportRows(schema: SchemaEntry | undefined, reportId?: string) {
  const { table, error, write } = useTable(schema?.excel ? schema : undefined);
  const rows = useMemo(() => table?.rows.filter((r) => String(r.report_id ?? "").trim().toLowerCase() === reportId?.toLowerCase()) ?? [], [table, reportId]);
  /** Adds a row for this report (report_id given) to the schema's Excel file. */
  const add = (values: Row) => write((t) => addRows(schema!.excel!, t, [{ ...values, report_id: reportId }]), { refetch: true, optimistic: (t) => withRows(t, [{ ...values, report_id: reportId }]) });
  /** Changes a row of this report (found in the whole table: the rows shown are its own objects). */
  const update = (row: Row, values: Row) =>
    write(
      (t) => updateRows(schema!.excel!, t, [{ index: t.rows.findIndex((r) => r.id === row.id), values }]),
      { optimistic: (t) => withUpdates(t, [{ index: t.rows.findIndex((r) => r.id === row.id), values }]) },
    );
  return { rows, error, ready: !!schema?.excel, add, update };
}

/**
 * Report viewer: a report in view mode with the AI on the left (it changes pages, filters, slicers and
 * reads the data), and tabs: the report and its guides (one per visual), read from the built-in Guides
 * schema (one Excel file for every report, rows by report_id).
 */
function Viewer() {
  const { report: reportId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/viewer" });
  useLast("viewer:report", reportId, (report) => navigate({ search: { report }, replace: true }));
  const entry = useCollection("reports").find((r) => r.id === reportId);
  const dataset = useCollection("datasets").find((d) => d.id === entry?.datasetId);
  const schemas = useCollection("schemas");
  const guides = useReportRows(schemas.find((s) => s.builtIn === "guides"), entry?.id);
  const { element, report, error, selection } = useEmbed(entry, "view");
  const pages = usePages(report);
  /** The left panel's tab: the assistant, or the guide of the visual clicked in the report. */
  const [panelTab, setPanelTab] = useState("assistant");
  const clicked = selection.page && selection.visual && pages.find((p) => p.name === selection.page)?.visuals.find((v) => v.name === selection.visual) ? { page: selection.page, visual: selection.visual } : undefined;
  const guidesOn = VIEWER_TABS.guides && guides.ready;
  /** The visual whose guide the panel shows: the last one clicked in the report, or picked in the panel. */
  const [chosen, setChosen] = useState<{ page: string; visual: string }>();
  // A click on a data visual shows its guide in the panel.
  useEffect(() => {
    if (!clicked) return;
    setChosen(clicked);
    if (guidesOn) setPanelTab("guide");
  }, [clicked?.page, clicked?.visual]); // eslint-disable-line react-hooks/exhaustive-deps
  const [conversationId, setConversationId] = useState<string>();
  const conversationExists = useCollection("conversations").some((c) => c.id === conversationId);
  useReportView(report, conversationId, conversationExists);

  const datasetRef = entry && { id: entry.datasetId, groupId: dataset?.groupId ?? entry.datasetGroupId };
  const agent: Agent | undefined = entry && {
    instructions: `Screen: the REPORT VIEWER. The user reads a Power BI report (view only: the report itself cannot change). You can change its page, filters, slicers and bookmarks, read the data of its visuals and run DAX on its model to answer precisely. Explain numbers simply.
Page state: the pages, the visuals of the CURRENT page (read a visual's data with get_visual_data) and the active filters, slicers and bookmarks. Other pages: list_pages (every page with its visuals) and read_page (a page in detail, without moving); go there with change_current_page, then act — the page state is read again after each change.
- Values: filter only on values that exist. A slicer lists its options in the page state; for any other column read the real values with run_dax. Translations and synonyms are fine (e.g. "automne" → the value "Fall").
- To filter, use the FIRST that fits: 1. a slicer of the current page on that field (update_slicer_value); 2. a page filter (update_page_filter); 3. a report filter (update_report_filter), when it must apply to every page; 4. a visual filter (update_visual_filter), only for one visual.
- Be fast. To change a filter, a slicer or the page: just do it (one call when possible) and say in one line what changed. Do NOT read the visuals' data (get_visual_data, read_page with data) unless the user asks about values or figures.
${DAX_RULES}`,
    tools: [...viewerTools(() => report, () => datasetRef, () => dataset?.info), daxTool(() => datasetRef), ...modelTools(() => dataset)],
    // Filters, slicers and pages are harmless and undone in a click: the agent acts at once (no plan, no question).
    direct: true,
    context: () => reportContext(entry, dataset),
    state: async () => (report ? `What the user sees now:\n${await pageText(report, false, datasetRef)}\n${await bookmarksText(report)}` : "The report is not loaded yet."),
  };

  const tabs = [
    {
      id: "report",
      label: "Report",
      icon: ChartPieIcon,
      content: (
        <>
          {error && <p className="px-6 py-2 text-[13px] text-destructive">{error}</p>}
          {!entry && <Empty icon={ChartPieIcon}>Choose a report, or add one from its Power BI link (menu at the top).</Empty>}
          <div ref={element} className={entry ? "min-h-0 flex-1" : "hidden"} />
        </>
      ),
    },
    ...(entry
      ? [
          { id: "dataset", label: "Dataset", icon: DatabaseIcon, content: <DatasetTab entry={entry} dataset={dataset} /> },
          { id: "context", label: "Context", icon: BookTextIcon, count: entry.contextIds.length, content: <ContextLinks what="report" ids={entry.contextIds} onChange={(contextIds) => void patch("reports", entry.id, { contextIds })} /> },
        ]
      : []),
  ];

  return (
    <Workspace
      title="Viewer"
      icon={ChartPieIcon}
      picker={<ReportPicker value={reportId} onChange={(report) => navigate({ search: { report } })} />}
      actions={
        <>
          <ReportActions entry={entry} report={report} />
          <Zoom report={report} />
        </>
      }
      ai={{
        scope: `viewer:${entry?.id}`,
        agent,
        placeholder: "Ask about this page, or to filter it…",
        empty: "Choose a report to talk about it.",
        onConversation: setConversationId,
        tab: panelTab,
        onTab: setPanelTab,
        tabs:
          entry && guidesOn
            ? [
                {
                  id: "guide",
                  label: "Guide",
                  icon: BookOpenIcon,
                  content: <GuidePanel report={report} reportRef={entry} entry={entry} dataset={dataset} pages={pages} open={chosen} onOpen={setChosen} guides={guides} />,
                },
              ]
            : undefined,
      }}
      tabs={tabs}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
    />
  );
}

/**
 * The report once loaded: fit to width; then the state the user left in the open AI conversation
 * (filters, slicers, page) is applied, and kept with it as it changes (a reload brings it back).
 */
function useReportView(report: Report | undefined, conversationId: string | undefined, exists: boolean) {
  useEffect(() => {
    if (!report) return;
    const fit = async () => {
      report.off("rendered", fit);
      const m = await models();
      await report.setZoom(1).catch(() => undefined);
      await report.updateSettings({ layoutType: m.LayoutType.Custom, customLayout: { displayOption: m.DisplayOption.FitToWidth } }).catch(() => undefined);
    };
    report.on("rendered", fit);
    return () => report.off("rendered", fit);
  }, [report]);
  useEffect(() => {
    if (!report || !conversationId) return;
    let live = true;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const save = () => {
      clearTimeout(timer);
      timer = setTimeout(async () => {
        const c = await find("conversations", conversationId);
        if (!live || !c) return; // a new conversation is saved with its first message
        const state = (await report.bookmarksManager.capture().catch(() => undefined))?.state;
        if (state && state !== c.viewState) await patch("conversations", conversationId, { viewState: state });
      }, 1500);
    };
    void find("conversations", conversationId).then(async (c) => {
      if (live && c?.viewState) await report.bookmarksManager.applyState(c.viewState).catch(() => undefined);
      if (live) save();
    });
    report.on("rendered", save);
    return () => {
      live = false;
      clearTimeout(timer);
      report.off("rendered", save);
    };
  }, [report, conversationId, exists]);
}

/* ---------------------------------- Guides --------------------------------- */

/** Visuals without data of their own (decoration, navigation): no guide, no preview, clicks ignored. */
const DECORATIONS = new Set(["textbox", "shape", "basicShape", "image", "actionButton", "pageNavigator", "bookmarkNavigator", "visualGroup", "group"]);
type VisualInfo = { name: string; title: string; type: string };
type PageInfo = { name: string; displayName: string; visuals: VisualInfo[] };

/** Pages of the embedded report with their data visuals, in reading order (top to bottom, then left to right). */
function usePages(report?: Report) {
  const [pages, setPages] = useState<PageInfo[]>([]);
  useEffect(() => {
    if (!report) return setPages([]);
    let live = true;
    (async () => {
      const list = (await report.getPages()).filter((p) => p.visibility !== 1);
      const out = await Promise.all(
        list.map(async (p) => ({
          name: p.name,
          displayName: p.displayName,
          visuals: (await p.getVisuals().catch(() => []))
            .filter((v) => v.layout.displayState?.mode !== 1 && !DECORATIONS.has(v.type))
            .sort((a, b) => Math.round((a.layout.y ?? 0) / 20) - Math.round((b.layout.y ?? 0) / 20) || (a.layout.x ?? 0) - (b.layout.x ?? 0))
            .map((v) => ({ name: v.name, title: v.title || v.type, type: v.type })),
        })),
      );
      if (live) setPages(out);
    })().catch(() => undefined);
    return () => void (live = false);
  }, [report]);
  return pages;
}

/**
 * The guide of one visual, in the assistant's panel: the visual (the one clicked in the report, or
 * picked here), its live preview, PDF / PNG / Explain, and its guide (editable with "Edit").
 */
function GuidePanel({ report, reportRef, entry, dataset, pages, open, onOpen, guides }: {
  report?: Report;
  reportRef: Ref;
  entry: ReportEntry;
  dataset?: DatasetEntry;
  pages: PageInfo[];
  open?: { page: string; visual: string };
  onOpen: (v: { page: string; visual: string }) => void;
  guides: { rows: Record<string, unknown>[]; add: (values: Row) => Promise<unknown>; update: (row: Row, values: Row) => Promise<unknown> };
}) {
  const preview = useRef<PreviewHandle>(undefined);
  const [context, setContext] = useState("");
  useEffect(() => void reportContext(entry, dataset).then(setContext), [entry, dataset]);
  const page = open && pages.find((p) => p.name === open.page);
  const visual = open && page?.visuals.find((v) => v.name === open.visual);
  const guide = open && guides.rows.find((r) => String(r.page ?? "") === open.page && String(r.visual ?? "") === open.visual);
  /** A visual's name for the user: its guide's title when it has one, else the visual's own title. */
  const guideOf = (p: string, v: string) => guides.rows.find((r) => String(r.page ?? "") === p && String(r.visual ?? "") === v);
  const nameOf = (p: string, v: string, fallback: string) => String(guideOf(p, v)?.title ?? "").trim() || fallback;
  const title = open ? nameOf(open.page, open.visual, visual?.title ?? open.visual) : "";
  return (
    <div className="flex flex-col gap-3 p-4">
      <select
        className="input h-9"
        aria-label="Visual"
        value={open ? `${open.page}/${open.visual}` : ""}
        onChange={(e) => {
          const [p, v] = e.target.value.split("/");
          if (p && v) onOpen({ page: p, visual: v });
        }}
      >
        <option value="">{pages.length ? "Choose a visual (or click one in the report)" : "Loading the report…"}</option>
        {pages.map((p) => (
          <optgroup key={p.name} label={p.displayName}>
            {p.visuals.map((v) => (
              <option key={v.name} value={`${p.name}/${v.name}`}>
                {nameOf(p.name, v.name, v.title)}
                {guideOf(p.name, v.name) ? " 📖" : ""}
              </option>
            ))}
          </optgroup>
        ))}
      </select>
      {!open ? (
        <Empty icon={BookOpenIcon}>Click a visual of the report, or choose one above, to read its guide.</Empty>
      ) : (
        // Keyed by visual: its preview, buttons and explanation are made again for each one (never stacked).
        <div key={`${open.page}/${open.visual}`} className="flex flex-col gap-3">
          <div>
            <p className="text-[11.5px] font-medium tracking-wide text-muted-foreground uppercase">{page?.displayName}</p>
            <h2 className="text-[16px] font-semibold">{title}</h2>
          </div>
          {VISUAL_PREVIEW && report && <VisualPreview report={report} reportRef={reportRef} page={open.page} visual={open.visual} handle={preview} />}
          {report && (EXPORT.visualPdf || EXPORT.visualPng || VISUAL_EXPLAIN) && <VisualActions report={report} page={open.page} visual={open.visual} title={title} preview={preview} context={context} dataset={dataset} />}
          <GuideCard guide={guide} ids={{ report: entry.id, ...open }} title={title} onAdd={guides.add} onUpdate={guides.update} />
        </div>
      )}
    </div>
  );
}

/** The report's semantic model: its tables (columns, measures) and relationships, as the AI reads it. */
function DatasetTab({ entry, dataset }: { entry: ReportEntry; dataset?: DatasetEntry }) {
  if (!dataset) return <Empty icon={DatabaseIcon}>The report's semantic model ({entry.datasetId}) is not in the app: add it in the Datasets screen to see it here.</Empty>;
  const info = dataset.info;
  return (
    <div className="mx-auto flex w-full max-w-5xl flex-col gap-4 overflow-y-auto p-6">
      <Section
        title={dataset.name}
        description={`Right: ${dataset.access ?? "unknown"}${dataset.error ? ` — ${dataset.error}` : ""}`}
        actions={
          <Button variant="outline" size="sm" asChild>
            <Link to="/datasets" search={{ dataset: dataset.id }}>
              <DatabaseIcon /> Open in Datasets
            </Link>
          </Button>
        }
      >
        {!info ? (
          <p className="text-[13px] text-muted-foreground">Structure not read.</p>
        ) : (
          <DataTable
            head={["Table", "Columns", "Measures"]}
            rows={info.tables
              .filter((t) => !t.hidden)
              .map((t) => [
                <span key="t" className="font-medium">{t.name}</span>,
                <span key="c" className="block text-[12.5px] whitespace-normal text-muted-foreground">{t.columns.filter((c) => !c.hidden).map((c) => c.name).join(", ")}</span>,
                <span key="m" className="block text-[12.5px] whitespace-normal">{t.measures.filter((m) => !m.hidden).map((m) => m.name).join(", ") || "—"}</span>,
              ])}
            empty="No table."
          />
        )}
      </Section>
      {info && info.relationships.length > 0 && (
        <Section title="Relationships">
          <ul className="flex flex-col gap-1 font-mono text-[12px]">
            {info.relationships.map((r, i) => (
              <li key={i}>
                {r.from} {r.cardinality} {r.to}
                {r.active ? "" : " (inactive)"}
              </li>
            ))}
          </ul>
        </Section>
      )}
    </div>
  );
}

/**
 * A visual's guide, read only; the pencil (guide editor on) shows its form, or — no guide yet — the
 * "Guide row" button and the visual's ids for the Guides schema.
 */
function GuideCard({ guide, ids, title, onAdd, onUpdate }: { guide?: Row; ids: { report: string; page: string; visual: string }; title: string; onAdd: (values: Row) => Promise<unknown>; onUpdate: (row: Row, values: Row) => Promise<unknown> }) {
  const [editing, setEditing] = useState(false);
  return (
    <div className="flex flex-col gap-3">
      <div className="flex flex-col gap-3 rounded-xl border bg-card p-4 shadow-xs">
        <div className="flex items-center gap-2">
          <span className="flex items-center gap-1.5 text-[12px] font-medium tracking-wide text-muted-foreground uppercase">
            <BookOpenIcon className="size-3.5" /> Guide
          </span>
          {GUIDE_EDITOR && (
            <Button variant={editing ? "secondary" : "outline"} size="sm" className="ml-auto h-7 text-[12px]" onClick={() => setEditing(!editing)}>
              <PencilIcon /> {editing ? "Done" : "Edit"}
            </Button>
          )}
        </div>
        {guide && editing ? (
          <GuideForm key={String(guide.id ?? `${ids.page}/${ids.visual}`)} guide={guide} onSave={(values) => onUpdate(guide, values)} />
        ) : guide ? (
          <div className="flex flex-col gap-2">
            {guide.content ? <Markdown>{String(guide.content)}</Markdown> : <p className="text-muted-foreground">This guide has no content yet.</p>}
          </div>
        ) : (
          <p className="text-[13.5px] text-muted-foreground">No guide for this visual yet.{GUIDE_EDITOR && !editing ? " Click Edit to add one." : ""}</p>
        )}
      </div>
      {GUIDE_EDITOR && editing && <VisualIds {...ids} onCreate={guide ? undefined : () => onAdd({ page: ids.page, visual: ids.visual, title: title.trim() || ids.visual })} />}
    </div>
  );
}

/** A guide written in the viewer (guide editor on): its title and content (markdown), saved on demand to the Excel file. */
function GuideForm({ guide, onSave }: { guide: Row; onSave: (values: Row) => Promise<unknown> }) {
  const saved = { title: String(guide.title ?? ""), content: String(guide.content ?? "") };
  const [draft, setDraft] = useState(saved);
  const [state, setState] = useState<{ busy?: boolean; error?: string; done?: boolean }>({});
  const changed = draft.title !== saved.title || draft.content !== saved.content;
  const save = () => {
    setState({ busy: true });
    onSave(draft).then(() => setState({ done: true }), (e) => setState({ error: describeError(e) }));
  };
  return (
    <div className="flex flex-col gap-3">
      <Input className="text-[15px] font-semibold" placeholder="Name shown for this visual" aria-label="Name shown for this visual (its guide's title)" value={draft.title} onChange={(e) => (setDraft({ ...draft, title: e.target.value }), setState({}))} />
      <RichText value={draft.content} onChange={(content) => (setDraft((d) => ({ ...d, content })), setState({}))} />
      <div className="flex items-center gap-2">
        <Button size="sm" disabled={!changed || state.busy || !draft.title.trim()} onClick={save}>
          <CheckIcon /> {state.busy ? "Saving…" : "Save"}
        </Button>
        {changed && (
          <Button size="sm" variant="ghost" onClick={() => setDraft(saved)}>
            Cancel
          </Button>
        )}
        {state.done && !changed && <span className="text-[12.5px] text-success">Saved in the Excel file.</span>}
        {state.error && <span className="text-[12.5px] break-words text-destructive">{state.error}</span>}
      </div>
    </div>
  );
}

/** The ids to write in the Guides schema for this visual (report_id, page, visual), each with its copy button. */
function VisualIds({ report, page, visual, onCreate }: { report: string; page: string; visual: string; onCreate?: () => Promise<unknown> }) {
  const [copied, setCopied] = useState<string>();
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  const create = () => {
    setState({ busy: true });
    // Shown in a toast as well: the card may be replaced by the new guide's form once the row exists.
    withToast("Adding the guide row…", onCreate!(), "Guide row added to the Excel file.").then(() => setState({}), (e) => setState({ error: describeError(e) }));
  };
  const row = (label: string, value: string) => (
    <div key={label} className="flex items-center gap-2">
      <span className="w-14 shrink-0 text-muted-foreground">{label}</span>
      <code className="min-w-0 flex-1 truncate rounded-md border bg-card px-2 py-1 font-mono text-[12px]" title={value}>
        {value}
      </code>
      <Button variant="ghost" size="icon" className="size-7" aria-label={`Copy the ${label} id`} onClick={() => void navigator.clipboard.writeText(value).then(() => setCopied(label))}>
        {copied === label ? <CheckIcon className="text-success" /> : <CopyIcon />}
      </Button>
    </div>
  );
  return (
    <div className="flex flex-col gap-1.5 rounded-xl bg-muted/60 p-3 text-[12.5px]">
      <span className="flex items-center gap-2 text-[12px] font-medium text-muted-foreground">
        Ids for the Guides schema
        {onCreate && (
          <Button variant="outline" size="sm" className="ml-auto h-7 text-[12px]" disabled={state.busy} title="Add this visual's row to the Guides Excel file (report, page, visual and title filled)" onClick={create}>
            <PlusIcon /> {state.busy ? "Adding…" : "Guide row"}
          </Button>
        )}
      </span>
      {row("report", report)}
      {row("page", page)}
      {row("visual", visual)}
      {state.error && <span className="text-[12px] break-words text-destructive">{state.error}</span>}
    </div>
  );
}

/** The user's filters on a page: report and page filters, and the slicers' selections (as page filters). */
async function filtersOf(report: Report, page: string, visual: string) {
  const current = (await report.getPages()).find((p) => p.name === page);
  if (!current) return { report: [], page: [] };
  const slicers = (await current.getVisuals()).filter((v) => v.type === "slicer" && v.name !== visual);
  const selections = (await Promise.all(slicers.map((v) => v.getSlicerState().then((st) => st.filters, () => [])))).flat();
  return { report: await report.getFilters().catch(() => []), page: [...(await current.getFilters().catch(() => [])), ...selections] };
}

/** A loaded preview: which visual (page/visual), its embed (to print) and its element (to capture). */
type PreviewHandle = { visual: string; embed: Awaited<ReturnType<typeof embedVisual>>; element: HTMLElement };

/** Preview height at most (px); the visual is scaled down to fit, never up. */
const PREVIEW_HEIGHT = 320;

/**
 * Live preview of the open visual: that visual alone in a second embed, drawn at its real size and
 * scaled down (ratio kept), with the report's filters and the slicers' selections, kept in step.
 */
function VisualPreview({ report, reportRef, page, visual, handle }: { report: Report; reportRef: Ref; page?: string; visual?: string; handle: { current?: PreviewHandle } }) {
  const element = useRef<HTMLDivElement>(null);
  const box = useRef<HTMLDivElement>(null);
  const embed = useRef<Awaited<ReturnType<typeof embedVisual>>>(undefined);
  const [size, setSize] = useState<{ width: number; height: number }>();
  const [width, setWidth] = useState(0);
  const [state, setState] = useState<{ loading?: boolean; error?: string }>({});
  const shown = !!(page && visual && size);
  useEffect(() => {
    if (!box.current) return;
    const observer = new ResizeObserver(([e]) => setWidth(e.contentRect.width));
    observer.observe(box.current);
    return () => observer.disconnect();
  }, []);
  useEffect(() => {
    if (!page || !visual || !element.current) return;
    let live = true;
    void (async () => {
      const target = (await (await report.getPages()).find((p) => p.name === page)?.getVisuals())?.find((v) => v.name === visual);
      if (!target || DECORATIONS.has(target.type)) return live && setSize(undefined);
      setSize({ width: target.layout.width || 400, height: target.layout.height || 300 });
      setState({ loading: true });
      handle.current = undefined;
      embed.current = await embedVisual(element.current!, reportRef, page, visual, await filtersOf(report, page, visual));
      if (!live) return;
      setState({});
      handle.current = { visual: `${page}/${visual}`, embed: embed.current, element: element.current! };
    })().catch((e) => live && setState({ error: describeError(e) }));
    return () => void (live = false);
  }, [report, reportRef.id, page, visual]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!page || !visual) return;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const follow = () => {
      clearTimeout(timer);
      timer = setTimeout(async () => embed.current && (await setVisualFilters(embed.current, await filtersOf(report, page, visual))), 800);
    };
    report.on("rendered", follow);
    return () => {
      clearTimeout(timer);
      report.off("rendered", follow);
    };
  }, [report, page, visual]);
  const scale = size && width ? Math.min(1, width / size.width, PREVIEW_HEIGHT / size.height) : 0;
  return (
    <div className={shown ? "rounded-xl border bg-card p-3 shadow-xs" : "hidden"}>
      <div ref={box} className="flex w-full justify-center">
        <div className="relative overflow-hidden rounded-lg" style={{ width: (size?.width ?? 0) * scale, height: (size?.height ?? 0) * scale }}>
          <div ref={element} style={{ width: size?.width, height: size?.height, transform: `scale(${scale})` }} className={cn("pointer-events-none absolute top-0 left-0 origin-top-left select-none", state.loading && "opacity-40")} />
        </div>
      </div>
      {state.error && <p className="text-[12px] break-words text-destructive">{state.error}</p>}
    </div>
  );
}

/** The open visual's buttons: PDF, PNG, Explain (the AI reads its data and explains it). */
function VisualActions({ report, page, visual, title, preview, context, dataset }: { report: Report; page: string; visual: string; title: string; preview: { current?: PreviewHandle }; context: string; dataset?: DatasetEntry }) {
  const [state, setState] = useState<{ busy?: "pdf" | "png" | "explain"; error?: string; explanation?: string }>({});
  const explain = () => {
    setState({ busy: "explain" });
    explainVisual(report, page, visual, context, dataset).then((explanation) => setState({ explanation }), (e) => setState({ error: describeError(e) }));
  };
  const run = (kind: "pdf" | "png") => {
    setState({ busy: kind });
    const ready = preview.current?.visual === `${page}/${visual}` ? preview.current : undefined;
    const task = kind === "pdf" ? (ready ? ready.embed.print() : printVisual(report, page, visual)) : ready ? capturePreviewPng(ready.element, title || visual) : captureVisualPng(report, page, visual, title);
    task.then(() => setState({}), (e) => setState({ error: describeError(e) }));
  };
  return (
    <div className="flex flex-col gap-2">
      <div className="flex items-center gap-2">
        {VISUAL_EXPLAIN && (
          <Button size="sm" disabled={!!state.busy} onClick={explain}>
            <SparklesIcon /> {state.busy === "explain" ? "Reading…" : "Explain"}
          </Button>
        )}
        {EXPORT.visualPdf && (
          <Button size="sm" variant="outline" disabled={!!state.busy} onClick={() => run("pdf")}>
            <DownloadIcon /> {state.busy === "pdf" ? "Preparing…" : "PDF"}
          </Button>
        )}
        {EXPORT.visualPng && (
          <Button size="sm" variant="outline" title="Captured from this tab: the browser asks once to share it" disabled={!!state.busy} onClick={() => run("png")}>
            <DownloadIcon /> {state.busy === "png" ? "Capturing…" : "PNG"}
          </Button>
        )}
      </div>
      {state.explanation && (
        <div className="rounded-xl border border-primary/20 bg-primary/5 p-4 text-[13.5px]">
          <Markdown>{state.explanation}</Markdown>
        </div>
      )}
      {state.error && <p className="text-[12px] break-words text-destructive">{state.error}</p>}
    </div>
  );
}
