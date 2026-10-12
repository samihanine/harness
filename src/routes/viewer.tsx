import { useEffect, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { icons, CheckIcon, ChevronRightIcon, FileTextIcon, LayoutPanelTopIcon, LinkIcon, PlusIcon, Settings2Icon, Trash2Icon, XIcon } from "lucide-react";
import { Drawer } from "@/components/drawer";
import { Markdown } from "@/components/markdown";
import { Chat } from "@/components/chat";
import { REPORT_VIEWER_EDITOR, VIEWER_TABS, visible } from "@/config";
import { RichText } from "@/components/rich-text";
import { COLORS, SWATCHES, isColor } from "@/lib/schema";
import { Picker } from "@/components/picker";
import { addReport, connectReportExcel } from "@/lib/library";
import { ReportActions } from "@/components/report-actions";
import { Zoom } from "@/components/zoom";
import { Empty, SidePanel } from "@/components/side-panel";
import { daxTool, viewTools } from "@/agent/pbi-tools";
import { modelTools } from "@/agent/model-tools";
import { addRows, deleteRows, updateRows, withRows, withUpdates, withoutRows, type Table } from "@/lib/excel";
import { useTable } from "@/lib/use-table";

type Write = ReturnType<typeof useTable>["write"];
import { errorText } from "@/lib/ms";
import { pageText, type Report } from "@/lib/pbi";
import { useCollection, type DatasetEntry, type ExcelEntry, type ReportEntry } from "@/lib/store";
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

  const agent = {
    instructions:
      "You help the user read a Power BI report. You see the page they are looking at (visuals with their data). You can change page, filters and slicers, and run DAX on the model to answer precisely. Explain numbers simply.",
    tools: [...viewTools(() => report), daxTool(() => entry && { id: entry.datasetId, groupId: entry.datasetGroupId }), ...modelTools(() => dataset)],
    context: async () => [libraryContext(entry, dataset), report ? `What the user sees now:\n${await pageText(report)}` : "The report is not loaded."].join("\n\n"),
  };

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={visible([
          { id: "info", label: "Info", content: <InfoTab excel={excels.find((e) => e.id === entry?.infoExcelId)} reportId={entry?.id} /> },
          { id: "guide", label: "Guides", content: <GuideTab excel={excels.find((e) => e.id === entry?.guidesExcelId)} reportId={entry?.id} report={report} selection={selection} /> },
          { id: "ai", label: "AI", content: entry ? <Chat scope={`viewer:${entry.id}`} agent={agent} placeholder="Ask about this page…" /> : <Empty>Choose a report.</Empty> },
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
    <div className="flex flex-col">
      <div className="flex h-10 items-center justify-between pr-2 pl-4">
        <span className="label">Links</span>
        <SettingsButton title="Edit the links" onClick={() => setEditing(true)} />
      </div>
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
                  className={`flex size-12 items-center justify-center rounded-[14px] bg-linear-to-b from-white/25 to-white/0 shadow-sm ring-1 ring-black/5 transition-transform duration-200 group-hover:scale-105 group-active:scale-95 [&_svg]:size-6 [&_svg]:stroke-[1.75] ${f.className}`}
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
                <button type="button" className="icon-btn" title="Delete" onClick={() => confirm("Delete this link?") && quiet(write(() => deleteRows(excel, [i]), { optimistic: (t) => withoutRows(t, [i]) }))}>
                  <Trash2Icon />
                </button>
              </div>
              {cell(i, "url", String(row.url ?? ""), "https://…")}
              <div className="flex items-center gap-1.5">
                <span className="w-24 shrink-0">{cell(i, "icon", String(row.icon ?? ""), "icon name")}</span>
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
      <p className="text-[11px] text-muted-foreground">Icons: lucide names (book-open, mail, video, calendar, file-text…).</p>
    </div>
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
function GuideTab({ excel, reportId, report, selection }: { excel?: ExcelEntry; reportId?: string; report?: Report; selection: Selection }) {
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
        <h3 className="text-[14px] font-medium">{String(row.title ?? "")}</h3>
        {row.content ? <Markdown>{String(row.content)}</Markdown> : <p className="text-[12px] text-muted-foreground">Empty guide.</p>}
        {links.length > 0 && (
          <div className="flex flex-col gap-1.5 border-t pt-3">
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
        {unmatched && (
          <p className="mx-3 mt-3 rounded-lg bg-muted/60 px-3 py-2 text-[12px] text-muted-foreground">
            No guide for <span className="font-medium text-foreground">{visualTitle(here)}</span> yet{REPORT_VIEWER_EDITOR ? ": write one with ⚙" : ""}.
          </p>
        )}
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
