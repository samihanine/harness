import { useEffect, useMemo, useRef, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { icons, BookOpenIcon, ChartColumnIcon, ChartPieIcon, CheckIcon, ChevronRightIcon, CopyIcon, DownloadIcon, LayoutPanelTopIcon, LinkIcon, PlusIcon, SparklesIcon } from "lucide-react";
import { Markdown } from "@/components/markdown";
import { ReportPicker } from "@/components/pickers";
import { ReportActions } from "@/components/report-actions";
import { RichText } from "@/components/rich-text";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Empty, Workspace } from "@/components/workspace";
import { Zoom } from "@/components/zoom";
import { EXPORT, GUIDE_EDITOR, VIEWER_TABS, VISUAL_EXPLAIN, VISUAL_PREVIEW } from "@/config";
import { reportContext } from "@/agent/context";
import { explainVisual } from "@/agent/explain";
import { describeError, type Agent } from "@/agent/loop";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES, bookmarksText, daxTool, viewerTools } from "@/agent/pbi-tools";
import { useLast } from "@/lib/last";
import { embedVisual, models, pageText, setVisualFilters, type Ref, type Report } from "@/lib/pbi";
import { SWATCHES, isColor, type Row } from "@/lib/schema";
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
 * Report viewer: a report in view mode with the AI on the left (it changes pages, filters, slicers and
 * reads the data), and tabs: the report, its guides (one per visual) and its links, both read from the
 * built-in schemas (one Excel file for every report, rows by report_id).
 */
function Viewer() {
  const { report: reportId, tab } = Route.useSearch();
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
    instructions: `Screen: the REPORT VIEWER. The user reads a Power BI report (view only: the report itself cannot change). You can change its page, filters, slicers and bookmarks, read the data of its visuals and run DAX on its model to answer precisely. Explain numbers simply.
Page state: the pages, the visuals of the current page (read a visual's data with get_visual_data) and the active filters, slicers and bookmarks.
- Values: filter only on values that exist. A slicer lists its options in the page state; for any other column read the real values with run_dax. Translations and synonyms are fine (e.g. "automne" → the value "Fall").
- To filter, use the FIRST that fits: 1. a slicer of the current page on that field (update_slicer_value); 2. a page filter (update_page_filter); 3. a report filter (update_report_filter), when it must apply to every page; 4. a visual filter (update_visual_filter), only for one visual.
- After a change, give what changed and the key figures (read with get_visual_data).
${DAX_RULES}`,
    tools: [...viewerTools(() => report, () => datasetRef, () => dataset?.info), daxTool(() => datasetRef), ...modelTools(() => dataset)],
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
    ...(VIEWER_TABS.guides && entry
      ? [
          {
            id: "guides",
            label: "Guides",
            icon: BookOpenIcon,
            count: guides.rows.length || undefined,
            content: guides.ready ? (
              <GuidesTab rows={guides.rows} onAdd={guides.add} onUpdate={guides.update} reportRef={entry} report={report} selection={selection} dataset={dataset} entry={entry} />
            ) : (
              <Empty icon={BookOpenIcon}>No guides file yet: set the Excel file of the built-in "Guides" schema in the Excel screen.</Empty>
            ),
          },
        ]
      : []),
    ...(VIEWER_TABS.links && entry ? [{ id: "links", label: "Links", icon: LinkIcon, count: links.rows.length || undefined, content: links.rows.length ? <LinksTab rows={links.rows} /> : <Empty icon={LinkIcon}>No link for this report (rows of the built-in "Links" schema, Excel screen).</Empty> }] : []),
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
      ai={{ scope: `viewer:${entry?.id}`, agent, placeholder: "Ask about this page, or to filter it…", empty: "Choose a report to talk about it.", onConversation: setConversationId }}
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

/* ---------------------------------- Links ---------------------------------- */

const IconOf = ({ name }: { name?: unknown }) => {
  const key = String(name ?? "").replace(/(^|[-_ ])(\w)/g, (_, __, c: string) => c.toUpperCase()) as keyof typeof icons;
  const Icon = icons[key] ?? LinkIcon;
  return <Icon className="size-4 shrink-0" />;
};

/** The report's links (rows of the Links schema): app-like tiles. Edited in the Excel screen. */
function LinksTab({ rows }: { rows: Record<string, unknown>[] }) {
  return (
    <div className="min-h-0 flex-1 overflow-y-auto p-8">
      <div className="mx-auto grid max-w-5xl grid-cols-[repeat(auto-fill,minmax(150px,1fr))] gap-4">
        {rows.map((row, i) => {
          const color = String(row.color ?? "") || "gray";
          return (
            <a key={i} href={String(row.url ?? "#")} target="_blank" rel="noreferrer" title={String(row.description || row.url || "")} className="group flex min-w-0 flex-col items-center gap-3 rounded-xl border bg-card px-3 py-5 text-center shadow-xs transition-all hover:-translate-y-0.5 hover:border-primary/40 hover:shadow-md">
              <span className={cn("flex size-14 items-center justify-center rounded-2xl text-white shadow-sm [&_svg]:size-7 [&_svg]:stroke-[1.75]", isColor(color) ? SWATCHES[color] : "bg-slate-400")}>
                <IconOf name={row.icon} />
              </span>
              <span className="line-clamp-2 w-full text-[13.5px] leading-tight font-semibold">{String(row.label || row.url || "")}</span>
              {row.description ? <span className="line-clamp-2 text-[12px] text-muted-foreground">{String(row.description)}</span> : null}
            </a>
          );
        })}
      </div>
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
 * The report's guides, one per visual: the pages and their data visuals on the left (a book on those
 * with a guide), the open visual on the right: its live preview, PDF / PNG / Explain, and its guide. A
 * click on a visual in the Report tab opens it here.
 */
function GuidesTab({ rows, onAdd, onUpdate, reportRef, report, selection, dataset, entry }: { rows: Record<string, unknown>[]; onAdd: (values: Row) => Promise<unknown>; onUpdate: (row: Row, values: Row) => Promise<unknown>; reportRef: Ref; report?: Report; selection: Selection; dataset?: DatasetEntry; entry: ReportEntry }) {
  const pages = usePages(report);
  const [open, setOpen] = useState<{ page: string; visual: string }>();
  const [folded, setFolded] = useState<string[]>([]);
  const preview = useRef<PreviewHandle>(undefined);
  const [context, setContext] = useState("");
  useEffect(() => void reportContext(entry, dataset).then(setContext), [entry, dataset]);
  // A click on a data visual of the report opens it here (guide or not); decorations are ignored.
  useEffect(() => {
    const { page, visual } = selection;
    if (page && visual && pages.find((p) => p.name === page)?.visuals.some((v) => v.name === visual)) setOpen({ page, visual });
  }, [selection.page, selection.visual, pages]); // eslint-disable-line react-hooks/exhaustive-deps
  const guideOf = (p: string, v: string) => rows.find((r) => String(r.page ?? "") === p && String(r.visual ?? "") === v);
  const page = open && pages.find((p) => p.name === open.page);
  const visual = open && page?.visuals.find((v) => v.name === open.visual);
  const guide = open && guideOf(open.page, open.visual);

  return (
    <div className="flex min-h-0 flex-1">
      <nav className="flex w-72 shrink-0 flex-col overflow-y-auto border-r bg-card p-2" aria-label="Pages and visuals">
        {pages.length === 0 && <p className="px-3 py-2 text-[13px] text-muted-foreground">Loading the report…</p>}
        {pages.map((p) => {
          const isFolded = folded.includes(p.name);
          const count = p.visuals.filter((v) => guideOf(p.name, v.name)).length;
          return (
            <div key={p.name} className="mb-1">
              <button type="button" onClick={() => setFolded(isFolded ? folded.filter((x) => x !== p.name) : [...folded, p.name])} className="flex h-9 w-full items-center gap-2 rounded-md px-2 text-left text-[13px] font-semibold hover:bg-accent">
                <ChevronRightIcon className={cn("size-3.5 shrink-0 text-muted-foreground transition-transform", !isFolded && "rotate-90")} />
                <LayoutPanelTopIcon className="size-4 shrink-0 text-muted-foreground" />
                <span className="truncate">{p.displayName}</span>
                <span className="ml-auto text-[11px] font-normal text-muted-foreground tabular-nums">{count}/{p.visuals.length}</span>
              </button>
              {!isFolded && (
                <ul className="ml-4 border-l pl-2">
                  {p.visuals.map((v) => {
                    const on = open?.page === p.name && open.visual === v.name;
                    return (
                      <li key={v.name}>
                        <button type="button" onClick={() => setOpen({ page: p.name, visual: v.name })} className={cn("flex h-8 w-full items-center gap-2 rounded-md px-2 text-left text-[13px]", on ? "bg-accent font-medium text-accent-foreground" : "hover:bg-muted")}>
                          <ChartColumnIcon className="size-3.5 shrink-0 text-muted-foreground" />
                          <span className="truncate">{v.title}</span>
                          {guideOf(p.name, v.name) && <BookOpenIcon className="ml-auto size-3.5 shrink-0 text-primary" aria-label="Has a guide" />}
                        </button>
                      </li>
                    );
                  })}
                  {p.visuals.length === 0 && <li className="px-2 py-1.5 text-[12px] text-muted-foreground">No data visual.</li>}
                </ul>
              )}
            </div>
          );
        })}
      </nav>
      <div className="min-w-0 flex-1 overflow-y-auto">
        {!open ? (
          <Empty icon={BookOpenIcon}>Choose a visual on the left, or click one in the Report tab, to read its guide.</Empty>
        ) : (
          <div className="mx-auto flex max-w-3xl flex-col gap-4 p-6">
            <div>
              <p className="text-[12px] font-medium tracking-wide text-muted-foreground uppercase">{page?.displayName}</p>
              <h2 className="text-[19px] font-semibold">{visual?.title ?? open.visual}</h2>
            </div>
            {VISUAL_PREVIEW && report && <VisualPreview report={report} reportRef={reportRef} page={open.page} visual={open.visual} handle={preview} />}
            {report && (EXPORT.visualPdf || EXPORT.visualPng || VISUAL_EXPLAIN) && <VisualActions key={`${open.page}/${open.visual}`} report={report} page={open.page} visual={open.visual} title={visual?.title ?? open.visual} preview={preview} context={context} dataset={dataset} />}
            <div className="rounded-xl border bg-card p-5 shadow-xs">
              {guide && GUIDE_EDITOR ? (
                <GuideForm key={String(guide.id ?? `${open.page}/${open.visual}`)} guide={guide} onSave={(values) => onUpdate(guide, values)} />
              ) : guide ? (
                <div className="flex flex-col gap-2">
                  {guide.title ? <h3 className="text-[16px] font-semibold">{String(guide.title)}</h3> : null}
                  {guide.content ? <Markdown>{String(guide.content)}</Markdown> : <p className="text-muted-foreground">This guide has no content yet.</p>}
                </div>
              ) : (
                <p className="text-muted-foreground">No guide for this visual yet.</p>
              )}
            </div>
            {GUIDE_EDITOR && <VisualIds report={entry.id} page={open.page} visual={open.visual} onCreate={guide ? undefined : () => onAdd({ page: open.page, visual: open.visual, title: visual?.title ?? open.visual })} />}
          </div>
        )}
      </div>
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
      <Input className="text-[15px] font-semibold" placeholder="Title" aria-label="Title of the guide" value={draft.title} onChange={(e) => (setDraft({ ...draft, title: e.target.value }), setState({}))} />
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
    onCreate!().then(() => setState({}), (e) => setState({ error: describeError(e) }));
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
