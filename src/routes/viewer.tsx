import { useCallback, useEffect, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { icons, CheckIcon, ChevronRightIcon, FileTextIcon, LayoutPanelTopIcon, LinkIcon, PencilIcon, PlusIcon, Settings2Icon, Trash2Icon, XIcon } from "lucide-react";
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

/** No Excel file yet for this tab: paste a link (or pick one of the library); missing columns are created. */
function ConnectExcel({ reportId, kind, children }: { reportId?: string; kind: "info" | "guides"; children: ReactNode }) {
  const excels = useCollection("excels");
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  const connect = (value: string) => {
    if (!reportId || !value) return;
    setState({ busy: true });
    connectReportExcel(reportId, kind, value).then(() => setState({}), (e) => setState({ error: errorText(e) }));
  };
  if (!reportId) return <Empty>Choose a report.</Empty>;
  return (
    <div className="flex flex-col gap-2 p-4">
      <p className="text-[12px] text-muted-foreground">{children}</p>
      {REPORT_VIEWER_EDITOR && (
        <>
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
          <p className="text-[11px] text-muted-foreground">Missing columns are added to the file (an empty file gets its table).</p>
        </>
      )}
      {state.error && <p className="text-[12px] text-destructive">{state.error}</p>}
    </div>
  );
}

type Link = { label: string; url: string; icon: string; color: string };
const NEW_LINK: Link = { label: "New link", url: "https://", icon: "link", color: "blue" };
/** Solid background of a link color (palette name; older hex colors are used as they are). */
const fill = (color: string) => ({ className: `text-white ${isColor(color) ? SWATCHES[color] : ""}`, style: isColor(color) ? undefined : { background: color || "#a1a1aa" } });

/** Links listed in an Excel table (label, url, icon = lucide name, color): app-like tiles; ⚙ edits them as a table. */
function InfoTab({ excel, reportId }: { excel?: ExcelEntry; reportId?: string }) {
  const { table, error, write } = useTable(excel);
  const [managing, setManaging] = useState(false);
  if (!excel)
    return (
      <ConnectExcel reportId={reportId} kind="info">
        No info links for this report yet. They live in an Excel file (columns label, url, icon, color).
      </ConnectExcel>
    );
  if (error && !table) return <Empty>{error}</Empty>;
  if (!table) return <Empty>Loading…</Empty>;
  return (
    <div className="flex flex-col gap-4 p-4">
      {error && <p className="text-[12px] text-destructive">{error}</p>}
      {managing ? (
        <LinksTable excel={excel} table={table} write={write} />
      ) : (
        <>
          {table.rows.length === 0 && <Empty>No link yet{REPORT_VIEWER_EDITOR ? ": add some with ⚙" : ""}.</Empty>}
          <div className="grid grid-cols-2 gap-3.5">
            {table.rows.map((row, i) => {
              const f = fill(String(row.color ?? "") || "gray");
              return (
                <a
                  key={i}
                  href={String(row.url ?? "#")}
                  target="_blank"
                  rel="noreferrer"
                  title={String(row.url ?? "")}
                  className={`flex aspect-[4/3] flex-col justify-between rounded-2xl p-4 shadow-sm transition hover:-translate-y-0.5 hover:shadow-md [&_svg]:size-9 [&_svg]:stroke-[1.75] ${f.className}`}
                  style={f.style}
                >
                  <IconOf name={row.icon} />
                  <span className="line-clamp-2 text-[17px] leading-tight font-semibold">{String(row.label || row.url || "")}</span>
                </a>
              );
            })}
          </div>
        </>
      )}
      {REPORT_VIEWER_EDITOR && (
        <div className="flex items-center justify-end border-t pt-2 gap-1">
          {managing && (
            <a href={excel.link} target="_blank" rel="noreferrer" className="mr-auto text-[11px] text-muted-foreground underline">
              Open the Excel file
            </a>
          )}
          <button type="button" className={`icon-btn ${managing ? "bg-muted text-foreground" : ""}`} title={managing ? "Done" : "Edit the links"} onClick={() => setManaging(!managing)}>
            {managing ? <CheckIcon /> : <Settings2Icon />}
          </button>
        </div>
      )}
    </div>
  );
}

/** The links as an editable table: each change is saved to the Excel file when the field is left. */
/** Changes are shown at once and sent to the Excel file in the background (see useTable). */
function LinksTable({ excel, table, write }: { excel: ExcelEntry; table: Table; write: Write }) {
  const quiet = (p: Promise<unknown>) => void p.catch(() => undefined); // errors are shown by the tab
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
                <span className="w-24 shrink-0">
                  {cell(i, "icon", String(row.icon ?? ""), "icon name")}
                </span>
                <span className="flex flex-wrap gap-1">
                  {COLORS.map((c) => (
                    <button key={c} type="button" title={c} onClick={() => c !== color && save(i, { color: c })} className={`size-5 rounded-full ${SWATCHES[c]} ${color === c ? "ring-2 ring-foreground ring-offset-1 ring-offset-card" : ""}`} />
                  ))}
                </span>
              </div>
            </div>
          );
        })}
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

/**
 * Guides in an Excel table (title, content in markdown, links = "pageName" or "pageName/visualName",
 * separated by ;). Home: the report pages, then every guide. A clicked visual opens its guide; without
 * one, an editor is offered to create or link one (others go back to the list).
 */
function GuideTab({ excel, reportId, report, selection }: { excel?: ExcelEntry; reportId?: string; report?: Report; selection: Selection }) {
  const { table, error, write } = useTable(excel);
  const pages = usePages(report);
  const [view, setView] = useState<GuideView>({ kind: "home" });
  const [editing, setEditing] = useState<{ title: string; content: string } | null>(null);
  const rows = table?.rows ?? [];
  const linksOf = (row?: Record<string, unknown>) => String(row?.links ?? "").split(/\s*;\s*/).filter(Boolean);
  const here = selection.visual ? `${selection.page}/${selection.visual}` : undefined;

  // Follows the visual clicked in the report.
  useEffect(() => {
    if (!selection.visual) {
      if (selection.page) setView((v) => (v.kind === "page" ? { kind: "page", page: selection.page! } : v));
      return;
    }
    const index = rows.findIndex((r) => linksOf(r).includes(`${selection.page}/${selection.visual}`));
    setEditing(null);
    setView(index >= 0 ? { kind: "guide", index } : REPORT_VIEWER_EDITOR ? { kind: "page", page: selection.page ?? "" } : { kind: "home" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selection.page, selection.visual]);

  if (!excel)
    return (
      <ConnectExcel reportId={reportId} kind="guides">
        No guides for this report yet. They live in an Excel file (columns title, content, links).
      </ConnectExcel>
    );
  if (error && !table) return <Empty>{error}</Empty>;
  if (!table) return <Empty>Loading…</Empty>;

  const quiet = (p: Promise<unknown>) => void p.catch(() => undefined);
  const update = (index: number, values: Record<string, unknown>) => {
    const changes = [{ index, values }];
    quiet(write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  const create = (links: string, title: string) => {
    const values = { title, content: "", links };
    const index = rows.length;
    quiet(write((t) => addRows(excel, t, [values]), { refetch: true, optimistic: (t) => withRows(t, [values]) }));
    setView({ kind: "guide", index });
    setEditing({ title, content: "" });
  };
  const pageTitle = (name?: string) => pages.find((p) => p.name === name)?.displayName ?? name ?? "";
  const visualTitle = (link: string) => {
    const [page, visual] = link.split("/");
    if (!visual) return pageTitle(page);
    return pages.find((p) => p.name === page)?.visuals.find((v) => v.name === visual)?.title ?? visual;
  };
  const goTo = (link: string) => void report?.getPages().then((ps) => ps.find((p) => p.name === link.split("/")[0])?.setActive());
  const onPage = (page: string) => rows.map((row, index) => ({ row, index, links: linksOf(row) })).filter((g) => g.links.some((l) => l === page || l.startsWith(`${page}/`)));

  const crumbs = (
    <div className="flex min-w-0 items-center gap-1 border-b px-3 py-2 text-[12px] text-muted-foreground">
      <button type="button" className="truncate hover:text-foreground" onClick={() => (setView({ kind: "home" }), setEditing(null))}>
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

  if (view.kind === "guide") {
    const row = rows[view.index];
    if (!row) return crumbs;
    const links = linksOf(row);
    const setLinks = (next: string[]) => update(view.index, { links: next.join("; ") });
    const candidates = pages.flatMap((p) => [{ value: p.name, label: `${p.displayName} (page)` }, ...p.visuals.map((v) => ({ value: `${p.name}/${v.name}`, label: `${p.displayName} › ${v.title}` }))]).filter((c) => !links.includes(c.value));
    return (
      <div className="flex flex-col">
        {crumbs}
        <div className="flex flex-col gap-3 p-3">
          {editing ? (
            <>
              <input autoFocus className="input font-medium" value={editing.title} onChange={(e) => setEditing({ ...editing, title: e.target.value })} />
              <RichText value={editing.content} onChange={(content) => setEditing((x) => x && { ...x, content })} />
              <div className="flex gap-2">
                <button type="button" className="btn-primary" onClick={() => (update(view.index, editing), setEditing(null))}>
                  Save
                </button>
                <button type="button" className="btn" onClick={() => setEditing(null)}>
                  Cancel
                </button>
              </div>
            </>
          ) : (
            <>
              <div className="flex items-start gap-2">
                <h3 className="flex-1 text-[14px] font-medium">{String(row.title ?? "")}</h3>
                {REPORT_VIEWER_EDITOR && (
                  <button type="button" className="icon-btn" title="Edit" onClick={() => setEditing({ title: String(row.title ?? ""), content: String(row.content ?? "") })}>
                    <PencilIcon />
                  </button>
                )}
              </div>
              {row.content ? <Markdown>{String(row.content)}</Markdown> : <p className="text-[12px] text-muted-foreground">Empty guide.</p>}
            </>
          )}
          <div className="flex flex-col gap-1.5 border-t pt-3">
            <span className="label">Linked to</span>
            <div className="flex flex-wrap gap-1">
              {links.map((l) => (
                <span key={l} className="inline-flex h-6 items-center gap-1 rounded-md border bg-card pl-2 text-[12px]">
                  <button type="button" onClick={() => goTo(l)} className="max-w-48 truncate hover:underline">
                    {visualTitle(l)}
                  </button>
                  {REPORT_VIEWER_EDITOR && (
                    <button type="button" aria-label="Unlink" onClick={() => setLinks(links.filter((x) => x !== l))} className="px-1 text-muted-foreground hover:text-foreground">
                      <XIcon className="size-3" />
                    </button>
                  )}
                </span>
              ))}
              {links.length === 0 && <span className="text-[12px] text-muted-foreground">Nothing.</span>}
            </div>
            {REPORT_VIEWER_EDITOR && here && !links.includes(here) && (
              <button type="button" className="btn h-7 w-fit" onClick={() => setLinks([...links, here])}>
                <LinkIcon /> Link the selected visual
              </button>
            )}
            {REPORT_VIEWER_EDITOR && candidates.length > 0 && (
              <select className="input" value="" onChange={(e) => e.target.value && setLinks([...links, e.target.value])}>
                <option value="">Link another page or visual…</option>
                {candidates.map((c) => (
                  <option key={c.value} value={c.value}>
                    {c.label}
                  </option>
                ))}
              </select>
            )}
          </div>
          {REPORT_VIEWER_EDITOR && (
            <button
              type="button"
              className="btn w-fit text-destructive"
              onClick={() => {
                if (!confirm("Delete this guide?")) return;
                const index = view.index;
                quiet(write(() => deleteRows(excel, [index]), { optimistic: (t) => withoutRows(t, [index]) }));
                setView({ kind: "home" });
              }}
            >
              <Trash2Icon /> Delete guide
            </button>
          )}
          {error && <p className="text-[12px] text-destructive">{error}</p>}
        </div>
      </div>
    );
  }

  if (view.kind === "page") {
    const guides = onPage(view.page);
    const unmatched = REPORT_VIEWER_EDITOR && here && selection.page === view.page && !rows.some((r) => linksOf(r).includes(here));
    return (
      <div className="flex flex-col">
        {crumbs}
        {unmatched && (
          <div className="m-3 flex flex-col gap-2 rounded-lg border bg-card p-3 shadow-sm">
            <p className="text-[12px]">
              No guide for <span className="font-medium">{visualTitle(here)}</span> yet.
            </p>
            <button type="button" className="btn-primary h-7 w-fit" onClick={() => create(here, `About ${visualTitle(here)}`)}>
              <PlusIcon /> New guide for this visual
            </button>
            {rows.length > 0 && (
              <select
                className="input"
                value=""
                onChange={(e) => {
                  const index = Number(e.target.value);
                  update(index, { links: [...linksOf(rows[index]), here].join("; ") });
                  setView({ kind: "guide", index });
                }}
              >
                <option value="">Or link it to a guide…</option>
                {rows.map((r, i) => (
                  <option key={i} value={i}>
                    {String(r.title ?? "")}
                  </option>
                ))}
              </select>
            )}
          </div>
        )}
        <ul className="p-1.5">
          {guides.map((g) => item(g.index, String(g.row.title ?? ""), `${g.links.length} item(s)`, () => setView({ kind: "guide", index: g.index })))}
        </ul>
        {guides.length === 0 && !unmatched && <Empty>No guide on this page.</Empty>}
        {REPORT_VIEWER_EDITOR && (
          <button type="button" className="btn m-3 w-fit" onClick={() => create(view.page, `${pageTitle(view.page)} guide`)}>
            <PlusIcon /> Page guide
          </button>
        )}
        {error && <p className="px-3 text-[12px] text-destructive">{error}</p>}
      </div>
    );
  }

  return (
    <div className="flex flex-col">
      {crumbs}
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
        {rows.length === 0 && <li className="px-2 py-1.5 text-[12px] text-muted-foreground">No guide yet{REPORT_VIEWER_EDITOR ? ": click a visual or open a page to write one." : "."}</li>}
      </ul>
      {error && <p className="px-3 text-[12px] text-destructive">{error}</p>}
    </div>
  );
}
