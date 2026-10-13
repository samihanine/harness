import { useEffect, useRef, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { icons, BookOpenIcon, ChartColumnIcon, ArrowDownIcon, ArrowUpIcon, CheckIcon, DownloadIcon, ChevronRightIcon, ChevronsUpDownIcon, FileTextIcon, LayoutPanelTopIcon, LinkIcon, PlusIcon, SearchIcon, Settings2Icon, SparklesIcon, Trash2Icon, XIcon } from "lucide-react";
import { Popover } from "@/components/popover";
import { Drawer } from "@/components/drawer";
import { Markdown } from "@/components/markdown";
import { Chat } from "@/components/chat";
import { EXPORT, REPORT_VIEWER_EDITOR, VIEWER_TABS, VISUAL_EXPLAIN, VISUAL_PREVIEW, visible } from "@/config";
import { explainVisual } from "@/agent/explain";
import { RichText } from "@/components/rich-text";
import { COLORS, SWATCHES, isColor, isIdField } from "@/lib/schema";
import { Picker } from "@/components/picker";
import { addReport, connectReportExcel } from "@/lib/library";
import { ReportActions } from "@/components/report-actions";
import { Zoom } from "@/components/zoom";
import { Empty, SidePanel } from "@/components/side-panel";
import { DAX_RULES, daxTool, viewTools } from "@/agent/pbi-tools";
import { modelTools } from "@/agent/model-tools";
import { addRows, deleteRows, updateRows, withRows, withUpdates, withoutRows, type Table } from "@/lib/excel";
import { useTable } from "@/lib/use-table";

type Write = ReturnType<typeof useTable>["write"];
import { errorText } from "@/lib/ms";
import { describeError } from "@/agent/loop";
import { embedVisual, models, pageText, setVisualFilters, type Ref, type Report } from "@/lib/pbi";
import { capturePreviewPng, captureVisualPng, printVisual } from "@/lib/visual-image";
import { find, patch, useCollection, type DatasetEntry, type ExcelEntry, type ReportEntry } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";
import { useLast } from "@/lib/last";

export const Route = createFileRoute("/viewer")({
  validateSearch: (s: Record<string, unknown>) => ({ report: (s.report as string) || undefined }),
  component: Viewer,
});

export function ReportPicker({ value, onChange }: { value?: string; onChange: (id: string) => void }) {
  const reports = useCollection("reports");
  return (
    <Picker
      className="my-2 ml-2 w-72"
      items={reports.map((r) => ({ id: r.id, label: r.name, hint: r.editable === false ? "view only" : undefined }))}
      value={value}
      onChange={onChange}
      placeholder="Choose a report…"
      add={{ label: "Add a report", placeholder: "Power BI report link", run: async (link) => (await addReport(link)).id }}
    />
  );
}

/** AI context shared by the viewer and the builder: report, model and their notes. */
export const libraryContext = (report?: ReportEntry, dataset?: DatasetEntry) =>
  [
    report && `Report "${report.name}"${report.context ? `\nNotes: ${report.context}` : ""}`,
    dataset && `Semantic model "${dataset.name}"${dataset.context ? `\nNotes: ${dataset.context}` : ""}\n${dataset.model ?? ""}`,
  ]
    .filter(Boolean)
    .join("\n\n");

function Viewer() {
  const { report: reportId } = Route.useSearch();
  const navigate = useNavigate({ from: "/viewer" });
  useLast("viewer:report", reportId, (report) => navigate({ search: { report }, replace: true }));
  const entry = useCollection("reports").find((r) => r.id === reportId);
  const dataset = useCollection("datasets").find((d) => d.id === entry?.datasetId);
  const excels = useCollection("excels");
  const { element, report, error, selection } = useEmbed(entry, "view");
  const [conversationId, setConversationId] = useState<string>();
  const conversationExists = useCollection("conversations").some((c) => c.id === conversationId);
  useReportView(report, conversationId, conversationExists);

  const datasetRef = entry && { id: entry.datasetId, groupId: entry.datasetGroupId };
  const agent = {
    instructions:
      `You help the user read and explore a Power BI report. You see the page they are looking at (visuals with their data, slicers with their options). You can change page, slicers and filters, and run DAX on the model to answer precisely. Explain numbers simply.
How to work:
- Do the whole request in ONE go: when it is clear, act with the tools then answer. Never ask for a confirmation when the meaning is obvious, also through a translation or a synonym (e.g. "automne" → the season value "Fall", "last year" → the previous year). Ask only when the request really has several equally likely meanings.
- Values: filter only on values that exist. A slicer lists its options in the page state; for any other column, or when unsure, read the real values first with run_dax (EVALUATE TOPN(50, VALUES('Table'[Column])), or a search with CONTAINSSTRING). The filter tools refuse unknown values and list the valid ones: then pick the right one and retry, without asking.
- Not sure about a fact (a value, a total, which column holds something)? Check it with run_dax or the model tools instead of guessing.
- To filter or select, use the FIRST of these that fits, in this order:
  1. a slicer of the current page on that field (set_slicer) — what the user sees and expects;
  2. a page filter on the column (set_filter, level "page");
  3. a report filter (set_filter, level "report"), when it must apply to every page;
  4. a visual filter (set_visual_filter), only when the user means one visual.
  Find the field in the slicers first (their fields and options), then in the model (search_model).
- After a change, read the new page state and answer with what changed and the key figures.
${DAX_RULES}`,
    tools: [...viewTools(() => report, () => datasetRef, () => dataset?.info), daxTool(() => datasetRef), ...modelTools(() => dataset)],
    context: async () => libraryContext(entry, dataset),
    state: async () => (report ? `What the user sees now:\n${await pageText(report, true, datasetRef)}` : "The report is not loaded."),
  };

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={visible([
          { id: "info", label: "Info", content: <InfoTab excel={excels.find((e) => e.id === entry?.infoExcelId)} reportId={entry?.id} /> },
          { id: "guide", label: "Guides", content: <GuideTab excel={excels.find((e) => e.id === entry?.guidesExcelId)} reportId={entry?.id} reportRef={entry} report={report} selection={selection} context={libraryContext(entry, dataset)} dataset={dataset} /> },
          { id: "ai", label: "AI", content: entry ? <Chat scope={`viewer:${entry.id}`} agent={agent} placeholder="Ask about this page…" onConversation={setConversationId} /> : <Empty>Choose a report.</Empty> },
        ], VIEWER_TABS)}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex h-12 shrink-0 items-center gap-2 border-b bg-card pr-2">
          <ReportPicker value={reportId} onChange={(report) => navigate({ search: { report } })} />
          <span className="ml-auto" />
          <ReportActions entry={entry} report={report} />
          <Zoom report={report} />
        </div>
        {error && <p className="px-3 text-destructive">{error}</p>}
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

const IconOf = ({ name }: { name?: unknown }) => {
  const key = String(name ?? "").replace(/(^|[-_ ])(\w)/g, (_, __, c: string) => c.toUpperCase()) as keyof typeof icons;
  const Icon = icons[key] ?? LinkIcon;
  return <Icon className="size-4 shrink-0" />;
};

/** Small ⚙ button of a tab: opens its drawer (none when the viewer is read only). */
const SettingsButton = ({ title, onClick }: { title: string; onClick: () => void }) =>
  REPORT_VIEWER_EDITOR ? (
    <button type="button" className="icon-btn" title={title} onClick={onClick}>
      <Settings2Icon />
    </button>
  ) : null;

/** No Excel file yet for this tab (in its drawer): paste a link (or pick one of the library); missing columns are created. */
function ConnectExcel({ reportId, kind }: { reportId: string; kind: "info" | "guides" }) {
  const excels = useCollection("excels");
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  const connect = (value: string) => {
    if (!value) return;
    setState({ busy: true });
    connectReportExcel(reportId, kind, value).then(() => setState({}), (e) => setState({ error: errorText(e) }));
  };
  return (
    <div className="flex flex-col gap-2">
      <p className="text-[12px] text-muted-foreground">
        {kind === "info" ? "The links live in an Excel file (columns label, url, icon, color)." : "The guides live in an Excel file (columns title, content, links)."} Missing columns are added to the file (an empty file gets its table).
      </p>
      <input className="input" placeholder="Sharing link of an .xlsx (SharePoint / OneDrive)" value={link} onChange={(e) => setLink(e.target.value)} onKeyDown={(e) => e.key === "Enter" && connect(link.trim())} />
      <button type="button" className="btn-primary w-fit" disabled={!link.trim() || state.busy} onClick={() => connect(link.trim())}>
        {state.busy ? "Preparing the file…" : "Use this Excel file"}
      </button>
      {excels.length > 0 && (
        <select className="input" value="" disabled={state.busy} onChange={(e) => connect(e.target.value)}>
          <option value="">Or one of the library…</option>
          {excels.map((x) => (
            <option key={x.id} value={x.id}>
              {x.name}
            </option>
          ))}
        </select>
      )}
      {state.error && <p className="text-[12px] text-destructive">{state.error}</p>}
    </div>
  );
}

/** Drawer content of a tab: the Excel file to connect, its loading state, or the editor. */
function TabDrawer({ title, excel, reportId, kind, table, error, onClose, children }: {
  title: string;
  excel?: ExcelEntry;
  reportId: string;
  kind: "info" | "guides";
  table?: Table;
  error?: string;
  onClose: () => void;
  children: () => ReactNode;
}) {
  return (
    <Drawer
      title={title}
      onClose={onClose}
      footer={excel && (
        <a href={excel.link} target="_blank" rel="noreferrer" className="text-[12px] text-muted-foreground underline">
          Open the Excel file
        </a>
      )}
    >
      {!excel ? <ConnectExcel reportId={reportId} kind={kind} /> : table ? children() : <Empty>{error ?? "Loading…"}</Empty>}
      {excel && table && error && <p className="text-[12px] text-destructive">{error}</p>}
    </Drawer>
  );
}

type Link = { label: string; url: string; icon: string; color: string };
const NEW_LINK: Link = { label: "New link", url: "https://", icon: "link", color: "blue" };
/** Solid background of a link color (palette name; older hex colors are used as they are). */
const fill = (color: string) => ({ className: `text-white ${isColor(color) ? SWATCHES[color] : ""}`, style: isColor(color) ? undefined : { backgroundColor: color || "#a1a1aa" } });

/** Links listed in an Excel table (label, url, icon = lucide name, color): app-like icons, three per row; ⚙ edits them in a drawer. */
function InfoTab({ excel, reportId }: { excel?: ExcelEntry; reportId?: string }) {
  const { table, error, write } = useTable(excel);
  const [editing, setEditing] = useState(false);
  if (!reportId) return <Empty>Choose a report.</Empty>;
  return (
    <div className="flex flex-1 flex-col pt-3">
      {!excel ? (
        <Empty>No info links for this report yet{REPORT_VIEWER_EDITOR ? ": add some with ⚙" : ""}.</Empty>
      ) : !table ? (
        <Empty>{error ?? "Loading…"}</Empty>
      ) : table.rows.length === 0 ? (
        <Empty>No link yet{REPORT_VIEWER_EDITOR ? ": add some with ⚙" : ""}.</Empty>
      ) : (
        <div className="grid grid-cols-3 gap-x-1.5 gap-y-3 px-3 pb-4">
          {table.rows.map((row, i) => {
            const f = fill(String(row.color ?? "") || "gray");
            return (
              <a
                key={i}
                href={String(row.url ?? "#")}
                target="_blank"
                rel="noreferrer"
                title={String(row.url ?? "")}
                className="group flex min-w-0 flex-col items-center gap-1.5 rounded-2xl px-1 py-2.5 text-center transition-colors hover:bg-muted/70"
              >
                <span
                  className={`flex size-14 items-center justify-center rounded-2xl bg-linear-to-b from-white/25 to-white/0 shadow-sm ring-1 ring-black/5 transition-transform duration-200 group-hover:scale-105 group-active:scale-95 [&_svg]:size-7 [&_svg]:stroke-[1.75] ${f.className}`}
                  style={f.style}
                >
                  <IconOf name={row.icon} />
                </span>
                <span className="line-clamp-2 w-full text-[12px] leading-tight font-medium text-foreground/85">{String(row.label || row.url || "")}</span>
              </a>
            );
          })}
        </div>
      )}
      {error && table && <p className="px-4 text-[12px] text-destructive">{error}</p>}
      <div className="sticky bottom-0 mt-auto flex justify-end p-2">
        <SettingsButton title="Edit the links" onClick={() => setEditing(true)} />
      </div>
      {editing && (
        <TabDrawer title="Info links" excel={excel} reportId={reportId} kind="info" table={table} error={error} onClose={() => setEditing(false)}>
          {() => <LinksTable excel={excel!} table={table!} write={write} />}
        </TabDrawer>
      )}
    </div>
  );
}

/** The links as an editable list: each change is shown at once and sent to the Excel file in the background (see useTable). */
function LinksTable({ excel, table, write }: { excel: ExcelEntry; table: Table; write: Write }) {
  const quiet = (p: Promise<unknown>) => void p.catch(() => undefined); // errors are shown by the drawer
  const save = (index: number, values: Partial<Link>) => {
    const changes = [{ index, values }];
    quiet(write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  /** The order is the Excel rows' order: two links trade their content (each row keeps its id). */
  const move = (index: number, to: number) => {
    const names = table.fields.filter((f) => !isIdField(f)).map((f) => f.name);
    const pick = (row: Record<string, unknown>) => Object.fromEntries(names.map((n) => [n, row[n] ?? null]));
    const changes = [
      { index, values: pick(table.rows[to]) },
      { index: to, values: pick(table.rows[index]) },
    ];
    quiet(write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  // A plain function (not a component): the inputs keep their focus when the table re-renders.
  const cell = (index: number, name: keyof Link, value: string, placeholder?: string) => (
    <input
      key={`${index}:${value}`}
      defaultValue={value}
      placeholder={placeholder}
      onBlur={(e) => e.target.value !== value && save(index, { [name]: e.target.value })}
      onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
      className="h-8 w-full min-w-0 rounded-md border border-transparent bg-transparent px-1.5 outline-none hover:border-border focus:border-ring focus:bg-card"
    />
  );
  return (
    <div className="flex flex-col gap-2">
      <div className="flex flex-col divide-y rounded-lg border">
        {table.rows.map((row, i) => {
          const color = String(row.color ?? "") || "gray";
          const f = fill(color);
          return (
            <div key={i} className="flex flex-col gap-1 p-2">
              <div className="flex items-center gap-1.5">
                <span className={`flex size-8 shrink-0 items-center justify-center rounded-lg [&_svg]:size-4 ${f.className}`} style={f.style}>
                  <IconOf name={row.icon} />
                </span>
                {cell(i, "label", String(row.label ?? ""), "Title")}
                <button type="button" className="icon-btn" title="Move up" disabled={i === 0} onClick={() => move(i, i - 1)}>
                  <ArrowUpIcon />
                </button>
                <button type="button" className="icon-btn" title="Move down" disabled={i === table.rows.length - 1} onClick={() => move(i, i + 1)}>
                  <ArrowDownIcon />
                </button>
                <button type="button" className="icon-btn" title="Delete" onClick={() => confirm("Delete this link?") && quiet(write(() => deleteRows(excel, [i]), { optimistic: (t) => withoutRows(t, [i]) }))}>
                  <Trash2Icon />
                </button>
              </div>
              {cell(i, "url", String(row.url ?? ""), "https://…")}
              <div className="flex items-center gap-1.5">
                <IconPicker value={String(row.icon ?? "")} onChange={(icon) => save(i, { icon })} />
                <span className="flex flex-wrap gap-1">
                  {COLORS.map((c) => (
                    <button key={c} type="button" title={c} onClick={() => c !== color && save(i, { color: c })} className={`size-5 rounded-full ${SWATCHES[c]} ${color === c ? "ring-2 ring-foreground ring-offset-1 ring-offset-card" : ""}`} />
                  ))}
                </span>
              </div>
            </div>
          );
        })}
        {table.rows.length === 0 && <p className="p-3 text-[12px] text-muted-foreground">No link yet.</p>}
      </div>
      <button type="button" className="btn w-fit" onClick={() => quiet(write((t) => addRows(excel, t, [NEW_LINK]), { refetch: true, optimistic: (t) => withRows(t, [NEW_LINK]) }))}>
        <PlusIcon /> Add a link
      </button>
    </div>
  );
}

/** Lucide icon names as stored in the Excel file (book-open), from the library's names (BookOpen). */
const ICON_NAMES = Object.keys(icons).map((k) => k.replace(/([a-z0-9])([A-Z])/g, "$1-$2").toLowerCase());

/** Icon chosen in a searchable list (with previews) rather than typed. */
function IconPicker({ value, onChange }: { value: string; onChange: (name: string) => void }) {
  const [search, setSearch] = useState("");
  const q = search.trim().toLowerCase().replace(/\s+/g, "-");
  const shown = (q ? ICON_NAMES.filter((n) => n.includes(q)) : ICON_NAMES).slice(0, 120);
  return (
    <Popover
      width={300}
      trigger={(open) => (
        <button type="button" onClick={() => (setSearch(""), open())} className="flex h-8 w-40 shrink-0 items-center gap-1.5 rounded-md border bg-card px-2 text-left text-[12px] hover:bg-muted/50 [&_svg]:size-4">
          <IconOf name={value} />
          <span className="min-w-0 flex-1 truncate">{value || "icon"}</span>
          <ChevronsUpDownIcon className="size-3.5! shrink-0 text-muted-foreground" />
        </button>
      )}
    >
      {(close) => (
        <div className="flex flex-col gap-1.5">
          <label className="flex items-center gap-2 border-b px-1 pb-1.5">
            <SearchIcon className="size-3.5 text-muted-foreground" />
            <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search an icon (mail, book, chart…)" className="h-7 flex-1 bg-transparent text-[12px] outline-none" />
          </label>
          <div className="grid max-h-64 grid-cols-6 gap-0.5 overflow-y-auto">
            {shown.map((n) => (
              <button key={n} type="button" title={n} onClick={() => (onChange(n), close())} className={`flex h-9 items-center justify-center rounded-md hover:bg-muted [&_svg]:size-4.5 ${n === value ? "bg-muted ring-1 ring-ring" : ""}`}>
                <IconOf name={n} />
              </button>
            ))}
          </div>
          {shown.length === 0 && <p className="px-1 text-[12px] text-muted-foreground">No icon found.</p>}
        </div>
      )}
    </Popover>
  );
}

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
/** The visual a guide is about ("pageName/visualName"). One visual per guide: older rows listing several keep the first. */
const linkOf = (row?: Record<string, unknown>) => String(row?.links ?? "").split(/\s*;\s*/).filter(Boolean)[0] ?? "";

/**
 * Guides, one per visual, in an Excel table (title, content in markdown, links = "pageName/visualName").
 * Pages, then a page's data visuals in reading order (a book icon on those with a guide), then a
 * visual: its preview, PDF / PNG / Explain, and its guide. A click on a data visual of the report opens
 * it here, guide or not. ⚙ writes the guide of the open visual.
 */
function GuideTab({ excel, reportId, reportRef, report, selection, context = "", dataset }: { excel?: ExcelEntry; reportId?: string; reportRef?: Ref; report?: Report; selection: Selection; context?: string; dataset?: DatasetEntry }) {
  const { table, error, write } = useTable(excel);
  const pages = usePages(report);
  const [view, setView] = useState<GuideView>({ kind: "pages" });
  const [editing, setEditing] = useState(false);
  /** The preview of the open visual, used by the downloads when it is ready. */
  const preview = useRef<PreviewHandle>(undefined);
  const rows = table?.rows ?? [];

  // A click on a data visual of the report opens it here (guide or not); decorations are ignored.
  useEffect(() => {
    const { page, visual } = selection;
    if (page && visual && pages.find((p) => p.name === page)?.visuals.some((v) => v.name === visual)) setView({ kind: "visual", page, visual });
  }, [selection.page, selection.visual, pages]); // eslint-disable-line react-hooks/exhaustive-deps

  if (!reportId) return <Empty>Choose a report.</Empty>;

  const page = view.kind === "pages" ? undefined : pages.find((p) => p.name === view.page);
  const visual = view.kind === "visual" ? page?.visuals.find((v) => v.name === view.visual) : undefined;
  const guideOf = (pageName: string, visualName: string) => rows.findIndex((r) => linkOf(r) === `${pageName}/${visualName}`);
  const openPage = (name: string) => {
    setView({ kind: "page", page: name });
    void report?.getPages().then((ps) => ps.find((p) => p.name === name)?.setActive());
  };
  const crumb = (label: string, onClick?: () => void) => (
    <>
      <ChevronRightIcon className="size-3 shrink-0" />
      {onClick ? (
        <button type="button" className="truncate hover:text-foreground" onClick={onClick}>
          {label}
        </button>
      ) : (
        <span className="truncate text-foreground">{label}</span>
      )}
    </>
  );
  const item = (key: string, label: string, onClick: () => void, icon: ReactNode, hint?: ReactNode) => (
    <li key={key}>
      <button type="button" onClick={onClick} className="flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left hover:bg-muted">
        {icon}
        <span className="truncate">{label}</span>
        <span className="ml-auto flex shrink-0 items-center text-[11px] text-muted-foreground">{hint}</span>
      </button>
    </li>
  );

  let body: ReactNode;
  if (view.kind === "pages")
    body = (
      <ul className="p-1.5">
        {pages.map((p) => item(p.name, p.displayName, () => openPage(p.name), <LayoutPanelTopIcon className="size-3.5 shrink-0 text-muted-foreground" />, `${p.visuals.filter((v) => guideOf(p.name, v.name) >= 0).length} guide(s)`))}
        {pages.length === 0 && <li className="px-2 py-1.5 text-[12px] text-muted-foreground">Loading the report…</li>}
      </ul>
    );
  else if (view.kind === "page")
    body = (
      <ul className="p-1.5">
        {page?.visuals.map((v) =>
          item(v.name, v.title, () => setView({ kind: "visual", page: page.name, visual: v.name }), <ChartColumnIcon className="size-3.5 shrink-0 text-muted-foreground" />, guideOf(page.name, v.name) >= 0 ? <BookOpenIcon className="size-3.5 text-blue-600" aria-label="Has a guide" /> : null),
        )}
        {page && page.visuals.length === 0 && <li className="px-2 py-1.5 text-[12px] text-muted-foreground">No data visual on this page.</li>}
      </ul>
    );
  else {
    const row = rows[guideOf(view.page, view.visual)];
    body = (
      <div className="p-3">
        {!excel ? (
          <p className="text-[12px] text-muted-foreground">No guides file for this report{REPORT_VIEWER_EDITOR ? ": connect one with ⚙" : ""}.</p>
        ) : !table ? (
          <p className="text-[12px] text-muted-foreground">{error ?? "Loading…"}</p>
        ) : row?.content ? (
          <Markdown>{String(row.content)}</Markdown>
        ) : (
          <p className="text-[12px] text-muted-foreground">No guide written for this visual{REPORT_VIEWER_EDITOR ? ": write one with ⚙" : ""}.</p>
        )}
      </div>
    );
  }

  return (
    <div className="flex flex-col">
      <div className="flex h-10 min-w-0 shrink-0 items-center gap-1 border-b pr-2 pl-3 text-[12px] text-muted-foreground">
        <button type="button" className="truncate hover:text-foreground" onClick={() => setView({ kind: "pages" })}>
          Guides
        </button>
        {page && crumb(page.displayName, view.kind === "visual" ? () => setView({ kind: "page", page: page.name }) : undefined)}
        {visual && crumb(visual.title)}
        <span className="ml-auto" />
        <SettingsButton title={view.kind === "visual" ? "Write the guide of this visual" : "Guides file and list"} onClick={() => setEditing(true)} />
      </div>
      {VISUAL_PREVIEW && report && reportRef && <VisualPreview report={report} reportRef={reportRef} page={view.kind === "visual" ? view.page : undefined} visual={view.kind === "visual" ? view.visual : undefined} handle={preview} />}
      {view.kind === "visual" && report && (EXPORT.visualPdf || EXPORT.visualPng || VISUAL_EXPLAIN) && (
        <VisualDownload key={`${view.page}/${view.visual}`} report={report} page={view.page} visual={view.visual} title={visual?.title ?? view.visual} preview={preview} context={context} dataset={dataset} />
      )}
      {body}
      {error && table && <p className="px-3 py-2 text-[12px] text-destructive">{error}</p>}
      {editing && (
        <TabDrawer title="Guides" excel={excel} reportId={reportId} kind="guides" table={table} error={error} onClose={() => setEditing(false)}>
          {() =>
            view.kind === "visual" ? (
              <GuideEditor key={`${view.page}/${view.visual}`} excel={excel!} rows={rows} write={write} link={`${view.page}/${view.visual}`} title={visual?.title ?? view.visual} />
            ) : (
              <GuideList
                rows={rows}
                pages={pages}
                onOpen={(link) => {
                  const [p, v] = link.split("/");
                  setView({ kind: "visual", page: p, visual: v });
                  setEditing(false);
                }}
              />
            )
          }
        </TabDrawer>
      )}
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
const PREVIEW_HEIGHT = 220;

/**
 * Live preview of the clicked visual: that visual alone in a second embed, drawn at its real size in
 * the report and scaled down (same look, ratio kept), with the report's filters and the slicers'
 * selections, kept in step when they change. Read only (no clicks). The frame is made once, on the
 * first click, then only reloaded with the next visual (much faster); it stays mounted, hidden, when
 * no visual is selected.
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
      // Decorations (shapes, images, buttons) are not worth a preview.
      if (!target || ["shape", "basicShape", "image", "actionButton"].includes(target.type)) return live && setSize(undefined);
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
    <div className={shown ? "border-b p-2" : "hidden"}>
      <div ref={box} className="flex w-full justify-center">
        <div className="relative overflow-hidden" style={{ width: (size?.width ?? 0) * scale, height: (size?.height ?? 0) * scale }}>
          <div
            ref={element}
            style={{ width: size?.width, height: size?.height, transform: `scale(${scale})` }}
            className={`pointer-events-none absolute top-0 left-0 origin-top-left select-none ${state.loading ? "opacity-40" : ""}`}
          />
        </div>
      </div>
      {state.error && <p className="text-[11px] break-words text-destructive">{state.error}</p>}
    </div>
  );
}

/** The clicked visual's buttons (guide or not): PDF, PNG, Explain, above its guide. */
function VisualDownload({ report, page, visual, title, preview, context, dataset }: { report: Report; page: string; visual: string; title: string; preview: { current?: PreviewHandle }; context: string; dataset?: DatasetEntry }) {
  const [state, setState] = useState<{ busy?: "pdf" | "png" | "explain"; error?: string; explanation?: string }>({});
  const explain = () => {
    setState({ busy: "explain" });
    explainVisual(report, page, visual, context, dataset).then((explanation) => setState({ explanation }), (e) => setState({ error: describeError(e) }));
  };
  const run = (kind: "pdf" | "png") => {
    setState({ busy: kind });
    // The preview when it shows this visual (the report is not touched), else the report itself.
    const ready = preview.current?.visual === `${page}/${visual}` ? preview.current : undefined;
    const task =
      kind === "pdf"
        ? ready
          ? ready.embed.print()
          : printVisual(report, page, visual)
        : ready
          ? capturePreviewPng(ready.element, title || visual)
          : captureVisualPng(report, page, visual, title);
    task.then(() => setState({}), (e) => setState({ error: describeError(e) }));
  };
  return (
    <div className="flex flex-col gap-1 border-b px-3 py-2">
      <div className="flex items-center gap-1.5">
        {EXPORT.visualPdf && (
          <button type="button" className="btn h-7 shrink-0" disabled={!!state.busy} onClick={() => run("pdf")}>
            <DownloadIcon /> {state.busy === "pdf" ? "Preparing…" : "PDF"}
          </button>
        )}
        {EXPORT.visualPng && (
          <button type="button" className="btn h-7 shrink-0" title="Captured from this tab: the browser asks once to share it" disabled={!!state.busy} onClick={() => run("png")}>
            <DownloadIcon /> {state.busy === "png" ? "Capturing…" : "PNG"}
          </button>
        )}
        {VISUAL_EXPLAIN && (
          <button type="button" className="btn h-7 shrink-0" title="The AI explains this visual from its data and the active filters" disabled={!!state.busy} onClick={explain}>
            <SparklesIcon /> {state.busy === "explain" ? "Reading…" : "Explain"}
          </button>
        )}
      </div>
      {state.explanation && <div className="rounded-md bg-muted/50 p-2 text-[12px]"><Markdown>{state.explanation}</Markdown></div>}
      {state.error && <p className="text-[11px] break-words text-destructive">{state.error}</p>}
    </div>
  );
}

/** Drawer, no visual open: every guide with its visual (opens it); guides are written from a visual. */
function GuideList({ rows, pages, onOpen }: { rows: Record<string, unknown>[]; pages: PageInfo[]; onOpen: (link: string) => void }) {
  const where = (link: string) => {
    const [p, v] = link.split("/");
    const page = pages.find((x) => x.name === p);
    const visual = page?.visuals.find((x) => x.name === v);
    return page && visual ? `${page.displayName} › ${visual.title}` : undefined;
  };
  return (
    <div className="flex flex-col gap-2">
      <p className="text-[12px] text-muted-foreground">One guide per visual. To write one, open the visual (click it in the report, or choose it in a page), then ⚙.</p>
      <ul className="flex flex-col divide-y rounded-lg border">
        {rows.map((r, i) => {
          const link = linkOf(r);
          const label = link.includes("/") ? where(link) : undefined;
          return (
            <li key={i}>
              <button type="button" disabled={!label} onClick={() => onOpen(link)} className="flex w-full flex-col items-start px-3 py-2 text-left hover:bg-muted disabled:opacity-60">
                <span className="font-medium">{String(r.title ?? "") || "Untitled"}</span>
                <span className="text-[11px] text-muted-foreground">{label ?? "Not linked to a visual of this report"}</span>
              </button>
            </li>
          );
        })}
        {rows.length === 0 && <li className="p-3 text-[12px] text-muted-foreground">No guide yet.</li>}
      </ul>
    </div>
  );
}

/** Drawer, a visual open: write its guide (created on demand), or delete it. */
function GuideEditor({ excel, rows, write, link, title }: { excel: ExcelEntry; rows: Record<string, unknown>[]; write: Write; link: string; title: string }) {
  const quiet = (p: Promise<unknown>) => void p.catch(() => undefined); // errors are shown by the drawer
  const index = rows.findIndex((r) => linkOf(r) === link);
  const row = rows[index];
  if (!row) {
    const values = { title, content: "", links: link };
    return (
      <div className="flex flex-col gap-2">
        <p className="text-[12px] text-muted-foreground">No guide for "{title}" yet.</p>
        <button type="button" className="btn-primary w-fit" onClick={() => quiet(write((t) => addRows(excel, t, [values]), { refetch: true, optimistic: (t) => withRows(t, [values]) }))}>
          <PlusIcon /> Write its guide
        </button>
      </div>
    );
  }
  return (
    <GuideForm
      key={index}
      row={row}
      onSave={(values) => {
        const changes = [{ index, values }];
        quiet(write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
      }}
      onDelete={() => {
        if (!confirm("Delete this guide?")) return;
        quiet(write(() => deleteRows(excel, [index]), { optimistic: (t) => withoutRows(t, [index]) }));
      }}
    />
  );
}

function GuideForm({ row, onSave, onDelete }: { row: Record<string, unknown>; onSave: (values: { title: string; content: string }) => void; onDelete: () => void }) {
  const [draft, setDraft] = useState({ title: String(row.title ?? ""), content: String(row.content ?? "") });
  const changed = draft.title !== String(row.title ?? "") || draft.content !== String(row.content ?? "");
  return (
    <div className="flex flex-col gap-3">
      <input className="input font-medium" placeholder="Title" value={draft.title} onChange={(e) => setDraft({ ...draft, title: e.target.value })} />
      <RichText value={draft.content} onChange={(content) => setDraft((d) => ({ ...d, content }))} />
      <div className="flex gap-2">
        <button type="button" className="btn-primary" disabled={!changed} onClick={() => onSave(draft)}>
          <CheckIcon /> Save
        </button>
        <button type="button" className="btn ml-auto text-destructive" onClick={onDelete}>
          <Trash2Icon /> Delete guide
        </button>
      </div>
    </div>
  );
}
