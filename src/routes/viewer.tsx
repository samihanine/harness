import { useEffect, useRef, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { icons, ArrowDownIcon, ArrowUpIcon, CheckIcon, DownloadIcon, ChevronRightIcon, ChevronsUpDownIcon, FileTextIcon, LayoutPanelTopIcon, LinkIcon, PlusIcon, SearchIcon, Settings2Icon, Trash2Icon, XIcon } from "lucide-react";
import { Popover } from "@/components/popover";
import { Drawer } from "@/components/drawer";
import { Markdown } from "@/components/markdown";
import { Chat } from "@/components/chat";
import { EXPORT, REPORT_VIEWER_EDITOR, VIEWER_TABS, VISUAL_PREVIEW, visible } from "@/config";
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
import { captureVisualPng, printVisual } from "@/lib/visual-image";
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
    tools: [...viewTools(() => report, () => datasetRef), daxTool(() => datasetRef), ...modelTools(() => dataset)],
    context: async () => libraryContext(entry, dataset),
    state: async () => (report ? `What the user sees now:\n${await pageText(report, true, datasetRef)}` : "The report is not loaded."),
  };

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={visible([
          { id: "info", label: "Info", content: <InfoTab excel={excels.find((e) => e.id === entry?.infoExcelId)} reportId={entry?.id} /> },
          { id: "guide", label: "Guides", content: <GuideTab excel={excels.find((e) => e.id === entry?.guidesExcelId)} reportId={entry?.id} reportRef={entry} report={report} selection={selection} /> },
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
type PageInfo = { name: string; displayName: string; visuals: { name: string; title: string }[] };

/** Pages of the embedded report with their visuals (names and titles), for the guides. */
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
            .filter((v) => v.layout.displayState?.mode !== 1 && !["shape", "basicShape", "image", "actionButton"].includes(v.type))
            .map((v) => ({ name: v.name, title: v.title || v.type })),
        })),
      );
      if (live) setPages(out);
    })().catch(() => undefined);
    return () => void (live = false);
  }, [report]);
  return pages;
}

type GuideView = { kind: "home" } | { kind: "page"; page: string } | { kind: "guide"; index: number };
const linksOf = (row?: Record<string, unknown>) => String(row?.links ?? "").split(/\s*;\s*/).filter(Boolean);

/**
 * Guides in an Excel table (title, content in markdown, links = "pageName" or "pageName/visualName",
 * separated by ;). Home: the report pages, then every guide. A clicked visual opens its guide. The tab
 * only reads: ⚙ opens the drawer where guides are written, linked and deleted.
 */
function GuideTab({ excel, reportId, reportRef, report, selection }: { excel?: ExcelEntry; reportId?: string; reportRef?: Ref; report?: Report; selection: Selection }) {
  const { table, error, write } = useTable(excel);
  const pages = usePages(report);
  const [view, setView] = useState<GuideView>({ kind: "home" });
  /** Drawer: open, with the guide being edited (undefined = none chosen). */
  const [editor, setEditor] = useState<{ index?: number } | null>(null);
  const rows = table?.rows ?? [];
  const here = selection.visual ? `${selection.page}/${selection.visual}` : undefined;

  // Follows the visual clicked in the report.
  useEffect(() => {
    if (!selection.visual) {
      if (selection.page) setView((v) => (v.kind === "page" ? { kind: "page", page: selection.page! } : v));
      return;
    }
    const index = rows.findIndex((r) => linksOf(r).includes(`${selection.page}/${selection.visual}`));
    setView(index >= 0 ? { kind: "guide", index } : { kind: "page", page: selection.page ?? "" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selection.page, selection.visual]);

  if (!reportId) return <Empty>Choose a report.</Empty>;

  const pageTitle = (name?: string) => pages.find((p) => p.name === name)?.displayName ?? name ?? "";
  const visualTitle = (link: string) => {
    const [page, visual] = link.split("/");
    if (!visual) return pageTitle(page);
    return pages.find((p) => p.name === page)?.visuals.find((v) => v.name === visual)?.title ?? visual;
  };
  const goTo = (link: string) => void report?.getPages().then((ps) => ps.find((p) => p.name === link.split("/")[0])?.setActive());
  const onPage = (page: string) => rows.map((row, index) => ({ row, index, links: linksOf(row) })).filter((g) => g.links.some((l) => l === page || l.startsWith(`${page}/`)));

  const crumbs = (
    <div className="flex h-10 min-w-0 shrink-0 items-center gap-1 border-b pr-2 pl-3 text-[12px] text-muted-foreground">
      <button type="button" className="truncate hover:text-foreground" onClick={() => setView({ kind: "home" })}>
        Guides
      </button>
      {view.kind === "page" && (
        <>
          <ChevronRightIcon className="size-3 shrink-0" />
          <span className="truncate text-foreground">{pageTitle(view.page)}</span>
        </>
      )}
      {view.kind === "guide" && (
        <>
          <ChevronRightIcon className="size-3 shrink-0" />
          <span className="truncate text-foreground">{String(rows[view.index]?.title ?? "")}</span>
        </>
      )}
      <span className="ml-auto" />
      <SettingsButton title="Edit the guides" onClick={() => setEditor({ index: view.kind === "guide" ? view.index : undefined })} />
    </div>
  );
  const item = (key: string | number, label: string, hint: string, onClick: () => void, icon = <FileTextIcon className="size-3.5 shrink-0 text-muted-foreground" />) => (
    <li key={key}>
      <button type="button" onClick={onClick} className="flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left hover:bg-muted">
        {icon}
        <span className="truncate">{label}</span>
        <span className="ml-auto shrink-0 text-[11px] text-muted-foreground">{hint}</span>
      </button>
    </li>
  );

  let body: ReactNode;
  if (!excel) body = <Empty>No guides for this report yet{REPORT_VIEWER_EDITOR ? ": write some with ⚙" : ""}.</Empty>;
  else if (!table) body = <Empty>{error ?? "Loading…"}</Empty>;
  else if (view.kind === "guide") {
    const row = rows[view.index];
    const links = linksOf(row);
    body = row && (
      <div className="flex flex-col gap-3 p-3">
        {row.content ? <Markdown>{String(row.content)}</Markdown> : null}
        {links.length > 0 && (
          <div className="hidden flex-col gap-1.5 border-t pt-3">
            <span className="label">Linked to</span>
            <div className="flex flex-wrap gap-1">
              {links.map((l) => (
                <button key={l} type="button" onClick={() => goTo(l)} className="inline-flex h-6 max-w-48 items-center truncate rounded-md border bg-card px-2 text-[12px] hover:bg-muted">
                  {visualTitle(l)}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  } else if (view.kind === "page") {
    const guides = onPage(view.page);
    const unmatched = here && selection.page === view.page && !rows.some((r) => linksOf(r).includes(here));
    body = (
      <>
        <ul className="p-1.5">{guides.map((g) => item(g.index, String(g.row.title ?? ""), `${g.links.length} item(s)`, () => setView({ kind: "guide", index: g.index })))}</ul>
        {guides.length === 0 && !unmatched && <Empty>No guide on this page.</Empty>}
      </>
    );
  } else
    body = (
      <>
        <p className="label px-3 pt-3">Pages</p>
        <ul className="p-1.5">
          {pages.map((p) =>
            item(p.name, p.displayName, `${onPage(p.name).length} guide(s)`, () => {
              setView({ kind: "page", page: p.name });
              goTo(p.name);
            }, <LayoutPanelTopIcon className="size-3.5 shrink-0 text-muted-foreground" />),
          )}
          {pages.length === 0 && <li className="px-2 py-1.5 text-[12px] text-muted-foreground">Loading the report…</li>}
        </ul>
        <p className="label border-t px-3 pt-3">All guides</p>
        <ul className="p-1.5">
          {rows.map((r, i) => item(i, String(r.title ?? ""), `${linksOf(r).length} item(s)`, () => setView({ kind: "guide", index: i })))}
          {rows.length === 0 && <li className="px-2 py-1.5 text-[12px] text-muted-foreground">No guide yet.</li>}
        </ul>
      </>
    );

  return (
    <div className="flex flex-col">
      {crumbs}
      {VISUAL_PREVIEW && report && reportRef && <VisualPreview report={report} reportRef={reportRef} page={selection.page} visual={selection.visual} />}
      {(EXPORT.visualPdf || EXPORT.visualPng) && here && report && selection.page && selection.visual && (
        <VisualDownload key={here} report={report} page={selection.page} visual={selection.visual} title={visualTitle(here)} />
      )}
      {body}
      {error && table && <p className="px-3 py-2 text-[12px] text-destructive">{error}</p>}
      {editor && (
        <TabDrawer title="Guides" excel={excel} reportId={reportId} kind="guides" table={table} error={error} onClose={() => setEditor(null)}>
          {() => (
            <GuidesEditor
              excel={excel!}
              rows={rows}
              write={write}
              index={editor.index}
              onSelect={(index) => setEditor({ index })}
              onDeleted={() => (setEditor({}), setView({ kind: "home" }))}
              pages={pages}
              here={here}
              page={selection.page}
              pageTitle={pageTitle}
              visualTitle={visualTitle}
            />
          )}
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

/** Preview height at most (px); the visual is scaled down to fit, never up. */
const PREVIEW_HEIGHT = 220;

/**
 * Live preview of the clicked visual: that visual alone in a second embed, drawn at its real size in
 * the report and scaled down (same look, ratio kept), with the report's filters and the slicers'
 * selections, kept in step when they change. Read only (no clicks). The frame is made once, on the
 * first click, then only reloaded with the next visual (much faster); it stays mounted, hidden, when
 * no visual is selected.
 */
function VisualPreview({ report, reportRef, page, visual }: { report: Report; reportRef: Ref; page?: string; visual?: string }) {
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
      embed.current = await embedVisual(element.current!, reportRef, page, visual, await filtersOf(report, page, visual));
      if (live) setState({});
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

/** The clicked visual's download buttons (guide or not), above its guide. */
function VisualDownload({ report, page, visual, title }: { report: Report; page: string; visual: string; title: string }) {
  const [state, setState] = useState<{ busy?: "pdf" | "png"; error?: string }>({});
  const run = (kind: "pdf" | "png") => {
    setState({ busy: kind });
    (kind === "pdf" ? printVisual(report, page, visual) : captureVisualPng(report, page, visual, title)).then(() => setState({}), (e) => setState({ error: describeError(e) }));
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
      </div>
      {state.error && <p className="text-[11px] break-words text-destructive">{state.error}</p>}
    </div>
  );
}

/** Drawer of the guides: pick or create one, then edit its title, content and links, or delete it. */
function GuidesEditor({ excel, rows, write, index, onSelect, onDeleted, pages, here, page, pageTitle, visualTitle }: {
  excel: ExcelEntry;
  rows: Record<string, unknown>[];
  write: Write;
  index?: number;
  onSelect: (index?: number) => void;
  onDeleted: () => void;
  pages: PageInfo[];
  here?: string;
  page?: string;
  pageTitle: (name?: string) => string;
  visualTitle: (link: string) => string;
}) {
  const quiet = (p: Promise<unknown>) => void p.catch(() => undefined); // errors are shown by the drawer
  const update = (i: number, values: Record<string, unknown>) => {
    const changes = [{ index: i, values }];
    quiet(write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  const create = (links: string, title: string) => {
    const values = { title, content: "", links };
    quiet(write((t) => addRows(excel, t, [values]), { refetch: true, optimistic: (t) => withRows(t, [values]) }));
    onSelect(rows.length);
  };
  const row = index === undefined ? undefined : rows[index];
  const hereFree = here && !rows.some((r) => linksOf(r).includes(here));
  return (
    <>
      <div className="flex flex-col gap-2">
        <select className="input" value={index ?? ""} onChange={(e) => onSelect(e.target.value === "" ? undefined : Number(e.target.value))}>
          <option value="">Choose a guide…</option>
          {rows.map((r, i) => (
            <option key={i} value={i}>
              {String(r.title ?? "") || "Untitled"}
            </option>
          ))}
        </select>
        <div className="flex flex-wrap gap-1.5">
          {here && hereFree && (
            <button type="button" className="btn h-7" onClick={() => create(here, `About ${visualTitle(here)}`)}>
              <PlusIcon /> Guide for the selected visual
            </button>
          )}
          {page && (
            <button type="button" className="btn h-7" onClick={() => create(page, `${pageTitle(page)} guide`)}>
              <PlusIcon /> Guide for this page
            </button>
          )}
          {!page && (
            <button type="button" className="btn h-7" onClick={() => create("", "New guide")}>
              <PlusIcon /> New guide
            </button>
          )}
        </div>
      </div>
      {row && index !== undefined && (
        <GuideForm
          key={index}
          row={row}
          onSave={(values) => update(index, values)}
          links={linksOf(row)}
          setLinks={(next) => update(index, { links: next.join("; ") })}
          here={here}
          candidates={pages.flatMap((p) => [{ value: p.name, label: `${p.displayName} (page)` }, ...p.visuals.map((v) => ({ value: `${p.name}/${v.name}`, label: `${p.displayName} › ${v.title}` }))])}
          visualTitle={visualTitle}
          onDelete={() => {
            if (!confirm("Delete this guide?")) return;
            quiet(write(() => deleteRows(excel, [index]), { optimistic: (t) => withoutRows(t, [index]) }));
            onDeleted();
          }}
        />
      )}
    </>
  );
}

function GuideForm({ row, onSave, links, setLinks, here, candidates, visualTitle, onDelete }: {
  row: Record<string, unknown>;
  onSave: (values: { title: string; content: string }) => void;
  links: string[];
  setLinks: (links: string[]) => void;
  here?: string;
  candidates: { value: string; label: string }[];
  visualTitle: (link: string) => string;
  onDelete: () => void;
}) {
  const [draft, setDraft] = useState({ title: String(row.title ?? ""), content: String(row.content ?? "") });
  const changed = draft.title !== String(row.title ?? "") || draft.content !== String(row.content ?? "");
  const others = candidates.filter((c) => !links.includes(c.value));
  return (
    <div className="flex flex-col gap-3 border-t pt-4">
      <input className="input font-medium" placeholder="Title" value={draft.title} onChange={(e) => setDraft({ ...draft, title: e.target.value })} />
      <RichText value={draft.content} onChange={(content) => setDraft((d) => ({ ...d, content }))} />
      <button type="button" className="btn-primary w-fit" disabled={!changed} onClick={() => onSave(draft)}>
        <CheckIcon /> Save
      </button>
      <div className="flex flex-col gap-1.5 border-t pt-3">
        <span className="label">Linked to</span>
        <div className="flex flex-wrap gap-1">
          {links.map((l) => (
            <span key={l} className="inline-flex h-6 items-center gap-1 rounded-md border bg-card pl-2 text-[12px]">
              <span className="max-w-48 truncate">{visualTitle(l)}</span>
              <button type="button" aria-label="Unlink" onClick={() => setLinks(links.filter((x) => x !== l))} className="px-1 text-muted-foreground hover:text-foreground">
                <XIcon className="size-3" />
              </button>
            </span>
          ))}
          {links.length === 0 && <span className="text-[12px] text-muted-foreground">Nothing.</span>}
        </div>
        {here && !links.includes(here) && (
          <button type="button" className="btn h-7 w-fit" onClick={() => setLinks([...links, here])}>
            <LinkIcon /> Link the selected visual
          </button>
        )}
        {others.length > 0 && (
          <select className="input" value="" onChange={(e) => e.target.value && setLinks([...links, e.target.value])}>
            <option value="">Link a page or visual…</option>
            {others.map((c) => (
              <option key={c.value} value={c.value}>
                {c.label}
              </option>
            ))}
          </select>
        )}
      </div>
      <button type="button" className="btn w-fit text-destructive" onClick={onDelete}>
        <Trash2Icon /> Delete guide
      </button>
    </div>
  );
}
