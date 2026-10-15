import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import { icons, BookOpenIcon, ChartColumnIcon, CheckIcon, PlusIcon, ChevronRightIcon, CopyIcon, DownloadIcon, LayoutPanelTopIcon, LinkIcon, SparklesIcon } from "lucide-react";
import { Chat } from "@/components/chat";
import { Markdown } from "@/components/markdown";
import { Picker } from "@/components/picker";
import { ReportActions } from "@/components/report-actions";
import { Empty, PageHeader, SidePanel } from "@/components/side-panel";
import { Zoom } from "@/components/zoom";
import { EXPORT, GUIDE_EDITOR, VIEWER_TABS, VISUAL_EXPLAIN, VISUAL_PREVIEW, visible } from "@/config";
import { withContexts } from "@/agent/context";
import { explainVisual } from "@/agent/explain";
import { describeError, type Agent } from "@/agent/loop";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES, bookmarksText, daxTool, viewerTools } from "@/agent/pbi-tools";
import { useLast } from "@/lib/last";
import { modelSummary } from "@/lib/model";
import { embedVisual, models, pageText, setVisualFilters, type Ref, type Report } from "@/lib/pbi";
import { SWATCHES, isColor, type Row } from "@/lib/schema";
import { addRows, updateRows, withRows, withUpdates } from "@/lib/excel";
import { RichText } from "@/components/rich-text";
import { find, patch, useCollection, type DatasetEntry, type ReportEntry, type SchemaEntry } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";
import { useTable } from "@/lib/use-table";
import { capturePreviewPng, captureVisualPng, printVisual } from "@/lib/visual-image";

export const Route = createFileRoute("/viewer")({
  validateSearch: (s: Record<string, unknown>) => ({ report: (s.report as string) || undefined }),
  component: Viewer,
});

/** Picker of the library's reports (title of the screen). */
export function ReportPicker({ value, onChange, editable = false }: { value?: string; onChange: (id: string) => void; editable?: boolean }) {
  const reports = useCollection("reports").filter((r) => !editable || r.access === "edit");
  return (
    <Picker
      items={reports.map((r) => ({ id: r.id, label: r.name, hint: r.access === "edit" ? undefined : r.access === "view" ? "view only" : undefined }))}
      value={value}
      onChange={onChange}
      placeholder="Choose a report…"
      empty={<>No {editable ? "editable " : ""}report: add one in the <Link to="/" search={{ tab: "reports" }} className="underline">Library</Link>.</>}
    />
  );
}

/** What the AI knows of a report and its model: their names, the user's rights, the model's structure, the linked contexts. */
export const reportContext = (report: ReportEntry, dataset?: DatasetEntry) =>
  withContexts(
    [
      `Report "${report.name}" (id ${report.id}). The user's right on it: ${report.access ?? "unknown"}${report.accessNote ? ` (${report.accessNote})` : ""}.`,
      dataset
        ? `Semantic model "${dataset.name}" (right: ${dataset.access ?? "unknown"}${dataset.error ? `; ${dataset.error}` : ""})\n${dataset.info ? modelSummary(dataset.info) : "Structure not read."}`
        : "The semantic model is not in the library: DAX and field checks are not available.",
    ],
    [report.contextIds, dataset?.contextIds],
  );

/** Rows of a built-in schema (guides, links) for one report: none until its Excel file is set. */
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
 * Report viewer: a report in view mode, with its guides (one per visual) and links read from the
 * built-in schemas (one Excel file for every report, rows by report_id: the tabs show only when the
 * report has some), and the AI, which changes the page, the filters and the slicers.
 */
function Viewer() {
  const { report: reportId } = Route.useSearch();
  const navigate = useNavigate({ from: "/viewer" });
  useLast("viewer:report", reportId, (report) => navigate({ search: { report }, replace: true }));
  const entry = useCollection("reports").find((r) => r.id === reportId);
  const dataset = useCollection("datasets").find((d) => d.id === entry?.datasetId);
  const schemas = useCollection("schemas");
  const guides = useReportRows(schemas.find((s) => s.builtIn === "guides"), entry?.id);
  const links = useReportRows(schemas.find((s) => s.builtIn === "links"), entry?.id);
  const { element, report, error, selection } = useEmbed(entry, "view");
  const [conversationId, setConversationId] = useState<string>();
  const conversationExists = useCollection("conversations").some((c) => c.id === conversationId);
  useReportView(report, conversationId, conversationExists);

  const datasetRef = entry && { id: entry.datasetId, groupId: dataset?.groupId ?? entry.datasetGroupId };
  const agent: Agent | undefined = entry && {
    instructions: `You help the user read and explore a Power BI report (view only: you cannot change the report itself). The page state lists the pages, the visuals of the current page (without their data: read a visual's data with get_visual_data) and the active filters, slicers and bookmarks. You change page, filters, slicers and bookmarks, and run DAX on the model to answer precisely. Explain numbers simply.
How to work:
- Do the whole request in ONE go: when it is clear, act with the tools then answer. Never ask for a confirmation when the meaning is obvious, also through a translation or a synonym (e.g. "automne" → the season value "Fall"). Ask only when the request really has several equally likely meanings.
- Values: filter only on values that exist. A slicer lists its options in the page state; for any other column, or when unsure, read the real values first with run_dax. The filter tools refuse unknown values and list the valid ones: then pick the right one and retry, without asking.
- To filter, use the FIRST of these that fits: 1. a slicer of the current page on that field (update_slicer_value); 2. a page filter (update_page_filter); 3. a report filter (update_report_filter), when it must apply to every page; 4. a visual filter (update_visual_filter), only when the user means one visual.
- After a change, answer with what changed and the key figures (read them with get_visual_data).
${DAX_RULES}`,
    tools: [...viewerTools(() => report, () => datasetRef, () => dataset?.info), daxTool(() => datasetRef), ...modelTools(() => dataset)],
    context: () => reportContext(entry, dataset),
    // The lists (pages, visuals of the current page, filters, bookmarks), read again before each model call.
    state: async () => (report ? `What the user sees now:\n${await pageText(report, false, datasetRef)}\n${await bookmarksText(report)}` : "The report is not loaded yet."),
  };

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={visible(
          [
            // Guides: as soon as their Excel file is set (a report without guide still gets the visuals' previews and buttons).
            ...(guides.ready && entry ? [{ id: "guides", label: "Guides", icon: BookOpenIcon, content: <GuidesTab rows={guides.rows} onAdd={guides.add} onUpdate={guides.update} reportRef={entry} report={report} selection={selection} dataset={dataset} entry={entry} /> }] : []),
            ...(links.rows.length ? [{ id: "links", label: "Links", icon: LinkIcon, content: <LinksTab rows={links.rows} /> }] : []),
            { id: "ai", label: "AI", icon: SparklesIcon, content: agent ? <Chat scope={`viewer:${entry!.id}`} agent={agent} placeholder="Ask about this page, or to filter it…" onConversation={setConversationId} /> : <Empty>Choose a report.</Empty> },
          ],
          VIEWER_TABS,
        )}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        <PageHeader
          screen="Report viewer"
          picker={<ReportPicker value={reportId} onChange={(report) => navigate({ search: { report } })} />}
          actions={
            <>
              <ReportActions entry={entry} report={report} />
              <Zoom report={report} />
            </>
          }
        />
        {error && <p className="px-6 py-2 text-destructive">{error}</p>}
        {!entry && <Empty>Choose a report (add reports in the Library).</Empty>}
        <div ref={element} className={entry ? "min-h-0 flex-1" : "hidden"} />
      </div>
    </div>
  );
}

/**
 * The report once loaded: fit to width; then the state the user left in the open AI conversation
 * (filters, slicers, page) is applied, and kept with it as it changes (a reload brings it back).
 */
function useReportView(report: Report | undefined, conversationId: string | undefined, exists: boolean) {
  useEffect(() => {
    if (!report) return;
    // On the first render: set before, the page was still shown whole (fit to page).
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

/* ---------------------------------- Links ---------------------------------- */

const IconOf = ({ name }: { name?: unknown }) => {
  const key = String(name ?? "").replace(/(^|[-_ ])(\w)/g, (_, __, c: string) => c.toUpperCase()) as keyof typeof icons;
  const Icon = icons[key] ?? LinkIcon;
  return <Icon className="size-4 shrink-0" />;
};

/** The report's links (rows of the Links schema): app-like icons, three per row. Edited in the Excel editor. */
function LinksTab({ rows }: { rows: Record<string, unknown>[] }) {
  return (
    <div className="grid grid-cols-3 gap-x-2 gap-y-4 p-4">
      {rows.map((row, i) => {
        const color = String(row.color ?? "") || "gray";
        return (
          <a key={i} href={String(row.url ?? "#")} target="_blank" rel="noreferrer" title={String(row.description || row.url || "")} className="group flex min-w-0 flex-col items-center gap-2 rounded-2xl px-1 py-3 text-center transition-colors hover:bg-muted/70">
            <span className={`flex size-14 items-center justify-center rounded-2xl text-white shadow-sm ring-1 ring-black/5 transition-transform duration-200 group-hover:scale-105 [&_svg]:size-7 [&_svg]:stroke-[1.75] ${isColor(color) ? SWATCHES[color] : "bg-zinc-400"}`}>
              <IconOf name={row.icon} />
            </span>
            <span className="line-clamp-2 w-full text-[13px] leading-tight font-medium">{String(row.label || row.url || "")}</span>
          </a>
        );
      })}
    </div>
  );
}

/* ---------------------------------- Guides --------------------------------- */

type Selection = { page?: string; visual?: string };
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
            // Same band of 20 px = same line: read left to right.
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

type GuideView = { kind: "pages" } | { kind: "page"; page: string } | { kind: "visual"; page: string; visual: string };

/**
 * The report's guides (rows of the Guides schema: page, visual, title, content), one per visual. Pages,
 * then a page's data visuals in reading order (a book on those with a guide), then a visual: its
 * preview, PDF / PNG / Explain, and its guide. A click on a data visual of the report opens it here.
 * Guides are written in the Excel editor (schema "Guides"); the visual's ids are shown to copy.
 */
function GuidesTab({ rows, onAdd, onUpdate, reportRef, report, selection, dataset, entry }: { rows: Record<string, unknown>[]; onAdd: (values: Row) => Promise<unknown>; onUpdate: (row: Row, values: Row) => Promise<unknown>; reportRef?: Ref; report?: Report; selection: Selection; dataset?: DatasetEntry; entry?: ReportEntry }) {
  const pages = usePages(report);
  const [view, setView] = useState<GuideView>({ kind: "pages" });
  /** The preview of the open visual, used by the downloads when it is ready. */
  const preview = useRef<PreviewHandle>(undefined);
  const [context, setContext] = useState("");
  useEffect(() => void (entry && reportContext(entry, dataset).then(setContext)), [entry, dataset]);
  // A click on a data visual of the report opens it here (guide or not); decorations are ignored.
  useEffect(() => {
    const { page, visual } = selection;
    if (page && visual && pages.find((p) => p.name === page)?.visuals.some((v) => v.name === visual)) setView({ kind: "visual", page, visual });
  }, [selection.page, selection.visual, pages]); // eslint-disable-line react-hooks/exhaustive-deps

  const page = view.kind === "pages" ? undefined : pages.find((p) => p.name === view.page);
  const visual = view.kind === "visual" ? page?.visuals.find((v) => v.name === view.visual) : undefined;
  const guideOf = (p: string, v: string) => rows.find((r) => String(r.page ?? "") === p && String(r.visual ?? "") === v);
  const openPage = (name: string) => {
    setView({ kind: "page", page: name });
    void report?.getPages().then((ps) => ps.find((p) => p.name === name)?.setActive());
  };
  const crumb = (label: string, onClick?: () => void) => (
    <>
      <ChevronRightIcon className="size-3.5 shrink-0" />
      {onClick ? (
        <button type="button" className="truncate hover:text-foreground" onClick={onClick}>
          {label}
        </button>
      ) : (
        <span className="truncate font-medium text-foreground">{label}</span>
      )}
    </>
  );
  const item = (key: string, label: string, onClick: () => void, icon: ReactNode, hint?: ReactNode) => (
    <li key={key}>
      <button type="button" onClick={onClick} className="flex h-11 w-full items-center gap-3 rounded-xl px-3 text-left hover:bg-muted">
        {icon}
        <span className="truncate">{label}</span>
        <span className="ml-auto flex shrink-0 items-center text-[12.5px] text-muted-foreground">{hint}</span>
      </button>
    </li>
  );

  let body: ReactNode;
  if (view.kind === "pages")
    body = (
      <ul className="p-2">
        {pages.map((p) => item(p.name, p.displayName, () => openPage(p.name), <LayoutPanelTopIcon className="size-4 shrink-0 text-muted-foreground" />, `${p.visuals.filter((v) => guideOf(p.name, v.name)).length} guide(s)`))}
        {pages.length === 0 && <li className="px-3 py-2 text-muted-foreground">Loading the report…</li>}
      </ul>
    );
  else if (view.kind === "page")
    body = (
      <ul className="p-2">
        {page?.visuals.map((v) => item(v.name, v.title, () => setView({ kind: "visual", page: page.name, visual: v.name }), <ChartColumnIcon className="size-4 shrink-0 text-muted-foreground" />, guideOf(page.name, v.name) ? <BookOpenIcon className="size-4 text-foreground" aria-label="Has a guide" /> : null))}
        {page && page.visuals.length === 0 && <li className="px-3 py-2 text-muted-foreground">No data visual on this page.</li>}
      </ul>
    );
  else {
    const guide = guideOf(view.page, view.visual);
    body = (
      <div className="flex flex-col gap-3 p-4">
        {guide && GUIDE_EDITOR ? (
          <GuideForm key={String(guide.id ?? `${view.page}/${view.visual}`)} guide={guide} onSave={(values) => onUpdate(guide, values)} />
        ) : guide ? (
          <>
            {guide.title ? <h3 className="text-[16px] font-semibold">{String(guide.title)}</h3> : null}
            {guide.content ? <Markdown>{String(guide.content)}</Markdown> : <p className="text-muted-foreground">This guide has no content yet.</p>}
          </>
        ) : (
          <p className="text-muted-foreground">No guide for this visual.</p>
        )}
        {entry && GUIDE_EDITOR && <VisualIds report={entry.id} page={view.page} visual={view.visual} onCreate={guide ? undefined : () => onAdd({ page: view.page, visual: view.visual, title: visual?.title ?? view.visual })} />}
      </div>
    );
  }

  return (
    <div className="flex flex-col">
      <div className="flex h-12 min-w-0 shrink-0 items-center gap-1.5 border-b px-4 text-[13.5px] text-muted-foreground">
        <button type="button" className="truncate hover:text-foreground" onClick={() => setView({ kind: "pages" })}>
          Pages
        </button>
        {page && crumb(page.displayName, view.kind === "visual" ? () => setView({ kind: "page", page: page.name }) : undefined)}
        {visual && crumb(visual.title)}
      </div>
      {VISUAL_PREVIEW && report && reportRef && <VisualPreview report={report} reportRef={reportRef} page={view.kind === "visual" ? view.page : undefined} visual={view.kind === "visual" ? view.visual : undefined} handle={preview} />}
      {view.kind === "visual" && report && (EXPORT.visualPdf || EXPORT.visualPng || VISUAL_EXPLAIN) && (
        <VisualActions key={`${view.page}/${view.visual}`} report={report} page={view.page} visual={view.visual} title={visual?.title ?? view.visual} preview={preview} context={context} dataset={dataset} />
      )}
      {body}
    </div>
  );
}

/** A guide written in the viewer (guide editor on): its title and its content (markdown), saved on demand to the Excel file. */
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
    <div className="flex flex-col gap-2.5">
      <input className="input font-semibold" placeholder="Title" value={draft.title} onChange={(e) => (setDraft({ ...draft, title: e.target.value }), setState({}))} />
      <RichText value={draft.content} onChange={(content) => (setDraft((d) => ({ ...d, content })), setState({}))} />
      <div className="flex items-center gap-2">
        <button type="button" className="btn-primary h-8" disabled={!changed || state.busy || !draft.title.trim()} onClick={save}>
          <CheckIcon /> {state.busy ? "Saving…" : "Save"}
        </button>
        {changed && (
          <button type="button" className="btn h-8" onClick={() => setDraft(saved)}>
            Cancel
          </button>
        )}
        {state.done && !changed && <span className="text-[12.5px] text-green-700">Saved in the Excel file.</span>}
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
    onCreate!().then(() => setState({}), (e) => setState({ error: describeError(e) }));
  };
  const row = (label: string, value: string) => (
    <div key={label} className="flex items-center gap-2">
      <span className="w-14 shrink-0 text-muted-foreground">{label}</span>
      <code className="min-w-0 flex-1 truncate rounded-md bg-card px-2 py-1 font-mono text-[12px]" title={value}>
        {value}
      </code>
      <button type="button" className="icon-btn size-7" title={`Copy the ${label} id`} onClick={() => void navigator.clipboard.writeText(value).then(() => setCopied(label))}>
        {copied === label ? <CheckIcon className="text-green-700" /> : <CopyIcon />}
      </button>
    </div>
  );
  return (
    <div className="mt-2 flex flex-col gap-1.5 rounded-xl bg-muted/60 p-3 text-[12.5px]">
      <span className="flex items-center gap-2 text-[12px] font-medium text-muted-foreground">
        Ids for the Guides schema
        {onCreate && (
          <button type="button" className="btn ml-auto h-7 rounded-full px-2.5 text-[12px]" disabled={state.busy} title="Add this visual's row to the Guides Excel file (report, page, visual and title filled): then write its content in the Excel editor" onClick={create}>
            <PlusIcon /> {state.busy ? "Adding…" : "Guide row"}
          </button>
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
const PREVIEW_HEIGHT = 240;

/**
 * Live preview of the clicked visual: that visual alone in a second embed, drawn at its real size and
 * scaled down (ratio kept), with the report's filters and the slicers' selections, kept in step when
 * they change. Read only. The frame is made once, then only reloaded with the next visual.
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
  // The user changes a filter or a slicer: the preview follows (filters only, no reload).
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
    <div className={shown ? "border-b p-3" : "hidden"}>
      <div ref={box} className="flex w-full justify-center">
        <div className="relative overflow-hidden rounded-lg" style={{ width: (size?.width ?? 0) * scale, height: (size?.height ?? 0) * scale }}>
          <div ref={element} style={{ width: size?.width, height: size?.height, transform: `scale(${scale})` }} className={`pointer-events-none absolute top-0 left-0 origin-top-left select-none ${state.loading ? "opacity-40" : ""}`} />
        </div>
      </div>
      {state.error && <p className="text-[12px] break-words text-destructive">{state.error}</p>}
    </div>
  );
}

/** The clicked visual's buttons (guide or not): PDF, PNG, Explain, above its guide. */
function VisualActions({ report, page, visual, title, preview, context, dataset }: { report: Report; page: string; visual: string; title: string; preview: { current?: PreviewHandle }; context: string; dataset?: DatasetEntry }) {
  const [state, setState] = useState<{ busy?: "pdf" | "png" | "explain"; error?: string; explanation?: string }>({});
  const explain = () => {
    setState({ busy: "explain" });
    explainVisual(report, page, visual, context, dataset).then((explanation) => setState({ explanation }), (e) => setState({ error: describeError(e) }));
  };
  const run = (kind: "pdf" | "png") => {
    setState({ busy: kind });
    // The preview when it shows this visual (the report is not touched), else the report itself.
    const ready = preview.current?.visual === `${page}/${visual}` ? preview.current : undefined;
    const task = kind === "pdf" ? (ready ? ready.embed.print() : printVisual(report, page, visual)) : ready ? capturePreviewPng(ready.element, title || visual) : captureVisualPng(report, page, visual, title);
    task.then(() => setState({}), (e) => setState({ error: describeError(e) }));
  };
  return (
    <div className="flex flex-col gap-2 border-b px-4 py-3">
      <div className="flex items-center gap-2">
        {EXPORT.visualPdf && (
          <button type="button" className="btn h-8 rounded-full" disabled={!!state.busy} onClick={() => run("pdf")}>
            <DownloadIcon /> {state.busy === "pdf" ? "Preparing…" : "PDF"}
          </button>
        )}
        {EXPORT.visualPng && (
          <button type="button" className="btn h-8 rounded-full" title="Captured from this tab: the browser asks once to share it" disabled={!!state.busy} onClick={() => run("png")}>
            <DownloadIcon /> {state.busy === "png" ? "Capturing…" : "PNG"}
          </button>
        )}
        {VISUAL_EXPLAIN && (
          <button type="button" className="btn h-8 rounded-full" title="The AI explains this visual from its data and the active filters" disabled={!!state.busy} onClick={explain}>
            <SparklesIcon /> {state.busy === "explain" ? "Reading…" : "Explain"}
          </button>
        )}
      </div>
      {state.explanation && (
        <div className="rounded-xl bg-muted/60 p-3 text-[13.5px]">
          <Markdown>{state.explanation}</Markdown>
        </div>
      )}
      {state.error && <p className="text-[12px] break-words text-destructive">{state.error}</p>}
    </div>
  );
}
