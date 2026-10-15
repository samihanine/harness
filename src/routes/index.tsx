import { useRef, useState, type ReactNode } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import { BookTextIcon, ChartPieIcon, CheckIcon, DatabaseIcon, DownloadIcon, ExternalLinkIcon, FileUpIcon, PencilIcon, PlusIcon, RefreshCwIcon, SearchIcon, ShapesIcon, Trash2Icon } from "lucide-react";
import { Drawer } from "@/components/drawer";
import { Badge } from "@/components/field-input";
import { Markdown } from "@/components/markdown";
import { Popover } from "@/components/popover";
import { PageHeader } from "@/components/side-panel";
import { refreshModel } from "@/server/ms";
import { addDataset, addReport, addSchema, createDatasetFromExcel, downloadEntry, importEntries, refreshDataset, refreshReport, setBuiltInExcel, type Kind } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { patch, put, remove, useCollection, type Access, type ContextEntry, type SchemaEntry } from "@/lib/store";
import type { Color } from "@/lib/schema";

type Tab = Kind;
export const Route = createFileRoute("/")({
  validateSearch: (s: Record<string, unknown>): { tab?: Tab } => ({ tab: ["reports", "datasets", "schemas", "contexts"].includes(s.tab as string) ? (s.tab as Tab) : undefined }),
  component: Library,
});

const TABS: { id: Tab; label: string; icon: ReactNode }[] = [
  { id: "reports", label: "Dashboards", icon: <ChartPieIcon /> },
  { id: "datasets", label: "Datasets", icon: <DatabaseIcon /> },
  { id: "schemas", label: "Schemas", icon: <ShapesIcon /> },
  { id: "contexts", label: "Contexts", icon: <BookTextIcon /> },
];

const ACCESS: Record<Access, { label: string; color: Color; hint: string }> = {
  edit: { label: "Edit", color: "green", hint: "Its definition can be read and changed" },
  build: { label: "Build", color: "blue", hint: "DAX queries only (no definition)" },
  view: { label: "View", color: "gray", hint: "View only" },
};
const AccessBadge = ({ access, note }: { access?: Access; note?: string }) =>
  access ? (
    <span title={note ?? ACCESS[access].hint}>
      <Badge color={ACCESS[access].color}>{ACCESS[access].label}</Badge>
    </span>
  ) : (
    <span className="text-muted-foreground">…</span>
  );

/**
 * The library (kept in this browser): dashboards (Power BI reports), datasets (semantic models), schemas
 * (an Excel table + optional form) and contexts (markdown files given to the AI with what they are
 * linked to). Every entry can be downloaded, and imported again.
 */
function Library() {
  const { tab = "reports" } = Route.useSearch();
  const navigate = useNavigate({ from: "/" });
  const [search, setSearch] = useState("");
  const [adding, setAdding] = useState(false);
  const [state, setState] = useState<{ busy?: string; error?: string; done?: string }>({});
  const importInput = useRef<HTMLInputElement>(null);
  const reports = useCollection("reports");
  const datasets = useCollection("datasets");
  const schemas = useCollection("schemas");
  const contexts = useCollection("contexts");
  const counts: Record<Tab, number> = { reports: reports.length, datasets: datasets.length, schemas: schemas.length, contexts: contexts.length };
  const run = (busy: string, task: Promise<unknown>, done?: string) => {
    setState({ busy });
    task.then(() => setState({ done }), (e) => setState({ error: errorText(e) }));
  };
  /** A name of the app's own: the report or model keeps its name in Power BI (empty: back to that name, read again). */
  const rename = (kind: "reports" | "datasets", entry: { id: string; name: string }) => {
    const name = prompt("Name in this app (empty: the name in Power BI)", entry.name);
    if (name === null) return;
    if (name.trim()) return void patch(kind, entry.id, { name: name.trim(), renamed: true });
    void patch(kind, entry.id, { renamed: false }).then(() => run("Reading…", kind === "reports" ? refreshReport(reports.find((r) => r.id === entry.id)!) : refreshDataset(datasets.find((d) => d.id === entry.id)!), "Name read again."));
  };
  const q = search.trim().toLowerCase();
  const match = (...texts: (string | undefined)[]) => !q || texts.some((t) => t?.toLowerCase().includes(q));

  return (
    <div className="flex h-full flex-col">
      <PageHeader
        screen="Library"
        actions={
          <>
            <input
              ref={importInput}
              type="file"
              multiple
              accept=".json,.md,.markdown,.txt"
              hidden
              onChange={(e) => {
                const files = e.target.files;
                if (files?.length) run("Importing…", importEntries(files).then((n) => setState({ done: `${n} entr${n > 1 ? "ies" : "y"} imported.` })));
                e.target.value = "";
              }}
            />
            <button type="button" className="btn" title="Entries downloaded before (JSON), or markdown contexts" onClick={() => importInput.current?.click()}>
              <FileUpIcon /> Import
            </button>
            <button type="button" className="btn-primary" onClick={() => setAdding(true)}>
              <PlusIcon /> Add {TABS.find((t) => t.id === tab)!.label.toLowerCase().replace(/s$/, "")}
            </button>
          </>
        }
      >
        {TABS.map((t) => (
          <button key={t.id} type="button" className="tab" data-active={t.id === tab} onClick={() => navigate({ search: { tab: t.id }, replace: true })}>
            {t.label}
            <span className="text-[12.5px] opacity-60">{counts[t.id]}</span>
          </button>
        ))}
        <span className="mx-2 h-6 border-l" />
        <label className="chip w-64 cursor-text bg-muted">
          <SearchIcon />
          <input className="min-w-0 flex-1 bg-transparent outline-none placeholder:text-muted-foreground" placeholder="Search" value={search} onChange={(e) => setSearch(e.target.value)} />
        </label>
        {state.busy && <span className="text-[13px] text-muted-foreground">{state.busy}</span>}
        {state.done && <span className="text-[13px] text-green-700">{state.done}</span>}
        {state.error && <span className="max-w-xl truncate text-[13px] text-destructive" title={state.error}>{state.error}</span>}
      </PageHeader>

      <div className="min-h-0 flex-1 overflow-auto p-6">
        {tab === "reports" && (
          <Grid
            head={["Name", "Dataset", "Access", "Contexts", ""]}
            empty="No dashboard yet: add a Power BI report link."
            rows={reports.filter((r) => match(r.name)).map((r) => [
              <Title key="n" icon={<ChartPieIcon />} title={r.name} hint={r.appId ? "Power BI app" : undefined} />,
              <span key="d" className="text-muted-foreground">{datasets.find((d) => d.id === r.datasetId)?.name ?? "…"}</span>,
              <AccessBadge key="a" access={r.access} note={r.accessNote} />,
              <ContextLinks key="c" ids={r.contextIds} contexts={contexts} onChange={(contextIds) => void patch("reports", r.id, { contextIds })} />,
              <Actions key="x">
                <Link to="/viewer" search={{ report: r.id }} className="btn h-8">View</Link>
                {r.access === "edit" && <Link to="/editor" search={{ report: r.id }} className="btn h-8">Edit</Link>}
                <IconAction title="Rename (in this app only)" onClick={() => rename("reports", r)} icon={<PencilIcon />} />
                <IconAction title="Read again (name, model, access)" onClick={() => run("Reading…", refreshReport(r), "Read again.")} icon={<RefreshCwIcon />} />
                <IconAction title="Download (JSON)" onClick={() => downloadEntry("reports", r)} icon={<DownloadIcon />} />
                <IconAction title="Remove from the library" danger onClick={() => confirm(`Remove "${r.name}" from the library?`) && void remove("reports", r.id)} icon={<Trash2Icon />} />
              </Actions>,
            ])}
          />
        )}
        {tab === "datasets" && (
          <Grid
            head={["Name", "Tables", "Access", "Contexts", ""]}
            empty="No dataset yet: add a semantic model link, or create one from Excel files."
            rows={datasets.filter((d) => match(d.name)).map((d) => [
              <Title key="n" icon={<DatabaseIcon />} title={d.name} hint={d.error} danger={!!d.error} />,
              <span key="t" className="text-muted-foreground">{d.info ? `${d.info.tables.filter((t) => !t.hidden).length} tables · ${d.info.tables.reduce((n, t) => n + t.measures.length, 0)} measures` : "—"}</span>,
              <AccessBadge key="a" access={d.access} />,
              <ContextLinks key="c" ids={d.contextIds} contexts={contexts} onChange={(contextIds) => void patch("datasets", d.id, { contextIds })} />,
              <Actions key="x">
                {d.access === "edit" && <Link to="/model" search={{ dataset: d.id }} className="btn h-8">Edit</Link>}
                <IconAction title="Refresh its data in Power BI (about 8 refreshes a day on Pro)" onClick={() => run("Refreshing the data…", refreshModel({ data: { datasetId: d.id, groupId: d.groupId } }).then((r) => { if (r.status !== "Completed") throw new Error(`Refresh ${r.status}${r.error ? `: ${r.error}` : ""}`); }), "Data refreshed.")} icon={<DatabaseIcon />} />
                <IconAction title="Rename (in this app only)" onClick={() => rename("datasets", d)} icon={<PencilIcon />} />
                <IconAction title="Read the structure again" onClick={() => run("Reading the model…", refreshDataset(d), "Read again.")} icon={<RefreshCwIcon />} />
                <IconAction title="Download (JSON)" onClick={() => downloadEntry("datasets", d)} icon={<DownloadIcon />} />
                <IconAction title="Remove from the library" danger onClick={() => confirm(`Remove "${d.name}" from the library?`) && void remove("datasets", d.id)} icon={<Trash2Icon />} />
              </Actions>,
            ])}
          />
        )}
        {tab === "schemas" && (
          <Grid
            head={["Title", "Excel", "Form", "Fields", "Contexts", ""]}
            empty="No schema yet: add an Excel table (and its form)."
            rows={schemas.filter((s) => match(s.title, s.excel?.name)).map((s) => [
              <Title key="n" icon={<ShapesIcon />} title={s.title} hint={s.builtIn ? "built-in" : undefined} />,
              s.excel ? (
                <a key="e" href={s.excel.webUrl ?? s.excel.link} target="_blank" rel="noreferrer" className="truncate hover:underline">
                  {s.excel.name} <span className="text-muted-foreground">› {s.excel.table}</span>
                </a>
              ) : (
                <BuiltInExcel key="e" schema={s} />
              ),
              s.form ? <a key="f" href={s.form.url} target="_blank" rel="noreferrer" className="truncate hover:underline">{s.form.title || "Form"}</a> : <span key="f" className="text-muted-foreground">—</span>,
              <span key="c" className="text-muted-foreground">{s.fields.length}</span>,
              <ContextLinks key="x" ids={s.contextIds} contexts={contexts} onChange={(contextIds) => void patch("schemas", s.id, { contextIds })} />,
              <Actions key="a">
                {!s.builtIn && <Link to="/schemas" search={{ schema: s.id }} className="btn h-8">Schema</Link>}
                {s.excel && <Link to="/excel" search={{ schema: s.id }} className="btn h-8">Data</Link>}
                <IconAction title="Download (JSON)" onClick={() => downloadEntry("schemas", s)} icon={<DownloadIcon />} />
                {!s.builtIn && <IconAction title="Remove from the library (the files are kept)" danger onClick={() => confirm(`Remove "${s.title}" from the library? (the Excel file and the form are kept)`) && void remove("schemas", s.id)} icon={<Trash2Icon />} />}
              </Actions>,
            ])}
          />
        )}
        {tab === "contexts" && <Contexts contexts={contexts.filter((c) => match(c.name, c.text))} />}
      </div>
      {adding && <AddDrawer tab={tab} onClose={() => setAdding(false)} />}
    </div>
  );
}

/* ---------------------------------- Table ---------------------------------- */

function Grid({ head, rows, empty }: { head: string[]; rows: ReactNode[][]; empty: string }) {
  if (!rows.length) return <p className="card p-10 text-center text-muted-foreground">{empty}</p>;
  return (
    <div className="card overflow-hidden">
      <table className="w-full border-collapse">
        <thead className="bg-muted/40">
          <tr>
            {head.map((h, i) => (
              <th key={i} className="th border-b">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((cells, i) => (
            <tr key={i} className="border-b last:border-b-0 hover:bg-muted/30">
              {cells.map((c, k) => (
                <td key={k} className="h-16 max-w-80 px-4 align-middle">
                  {c}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const Title = ({ icon, title, hint, danger }: { icon: ReactNode; title: string; hint?: string; danger?: boolean }) => (
  <span className="flex min-w-0 items-center gap-3">
    <span className="flex size-9 shrink-0 items-center justify-center rounded-full bg-muted text-muted-foreground [&_svg]:size-4">{icon}</span>
    <span className="min-w-0">
      <span className="block truncate font-medium">{title}</span>
      {hint && <span className={`block truncate text-[12.5px] ${danger ? "text-destructive" : "text-muted-foreground"}`} title={hint}>{hint}</span>}
    </span>
  </span>
);

const Actions = ({ children }: { children: ReactNode }) => <span className="flex items-center justify-end gap-1">{children}</span>;
const IconAction = ({ title, icon, onClick, danger }: { title: string; icon: ReactNode; onClick: () => void; danger?: boolean }) => (
  <button type="button" title={title} className={`icon-btn ${danger ? "hover:text-destructive" : ""}`} onClick={onClick}>
    {icon}
  </button>
);

/** The contexts linked to an entry, chosen in a list. */
function ContextLinks({ ids = [], contexts, onChange }: { ids?: string[]; contexts: ContextEntry[]; onChange: (ids: string[]) => void }) {
  const linked = contexts.filter((c) => ids.includes(c.id));
  return (
    <Popover
      width={300}
      trigger={(open) => (
        <button type="button" onClick={open} className="flex max-w-60 items-center gap-1.5 rounded-lg px-1.5 py-1 text-left hover:bg-muted">
          {linked.length ? linked.map((c) => <Badge key={c.id} color="violet">{c.name}</Badge>) : <span className="text-muted-foreground">+ Link</span>}
        </button>
      )}
    >
      {() => (
        <div className="flex max-h-72 flex-col overflow-y-auto">
          {contexts.map((c) => {
            const on = ids.includes(c.id);
            return (
              <button key={c.id} type="button" onClick={() => onChange(on ? ids.filter((x) => x !== c.id) : [...ids, c.id])} className="flex items-center gap-2 rounded-lg px-2 py-1.5 text-left hover:bg-muted">
                <CheckIcon className={`size-4 ${on ? "" : "invisible"}`} />
                <span className="truncate">{c.name}</span>
              </button>
            );
          })}
          {!contexts.length && <p className="p-2 text-[13px] text-muted-foreground">No context yet: add markdown files in the Contexts tab.</p>}
        </div>
      )}
    </Popover>
  );
}

/** A built-in schema without Excel file: give one (missing columns are added to it). */
function BuiltInExcel({ schema }: { schema: SchemaEntry }) {
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: string; error?: string }>({});
  return (
    <Popover
      width={380}
      trigger={(open) => (
        <button type="button" className="btn h-8" onClick={open}>
          <PlusIcon /> Set the Excel file
        </button>
      )}
    >
      {(close) => (
        <div className="flex flex-col gap-2 p-1">
          <p className="text-[13px] text-muted-foreground">One Excel file for every report (column report_id). A file without table gets one; missing columns are added.</p>
          <input autoFocus className="input" placeholder="Sharing link of an .xlsx" value={link} onChange={(e) => setLink(e.target.value)} />
          <button
            type="button"
            className="btn-primary"
            disabled={!link.trim() || !!state.busy}
            onClick={() => {
              setState({ busy: "Preparing the file…" });
              setBuiltInExcel(schema, link.trim(), (step, done, total) => setState({ busy: `${step} (${done}/${total})` })).then(() => (setState({}), close()), (e) => setState({ error: errorText(e) }));
            }}
          >
            {state.busy ?? "Use this file"}
          </button>
          {state.error && <p className="text-[12px] break-words text-destructive">{state.error}</p>}
        </div>
      )}
    </Popover>
  );
}

/* --------------------------------- Contexts -------------------------------- */

function Contexts({ contexts }: { contexts: ContextEntry[] }) {
  const reports = useCollection("reports");
  const datasets = useCollection("datasets");
  const schemas = useCollection("schemas");
  const [open, setOpen] = useState<ContextEntry>();
  const uses = (id: string) => [...reports, ...datasets, ...schemas].filter((e) => e.contextIds?.includes(id)).length;
  return (
    <>
      <Grid
        head={["Name", "Size", "Linked to", "Updated", ""]}
        empty="No context yet: write one (company vocabulary, rules, documentation…). The AI receives it with the entries it is linked to."
        rows={contexts.map((c) => [
          <button key="n" type="button" className="text-left" onClick={() => setOpen(c)}>
            <Title icon={<BookTextIcon />} title={c.name} />
          </button>,
          <span key="s" className="text-muted-foreground">{(c.text.length / 1000).toFixed(1)} k characters</span>,
          <span key="l" className="text-muted-foreground">{uses(c.id)} entr{uses(c.id) === 1 ? "y" : "ies"}</span>,
          <span key="u" className="text-muted-foreground">{new Date(c.updatedAt).toLocaleDateString()}</span>,
          <Actions key="x">
            <button type="button" className="btn h-8" onClick={() => setOpen(c)}>Open</button>
            <IconAction title="Download (.md)" onClick={() => downloadEntry("contexts", c)} icon={<DownloadIcon />} />
            <IconAction title="Delete" danger onClick={() => confirm(`Delete the context "${c.name}"?`) && void remove("contexts", c.id)} icon={<Trash2Icon />} />
          </Actions>,
        ])}
      />
      {open && <ContextDrawer context={open} onClose={() => setOpen(undefined)} />}
    </>
  );
}

function ContextDrawer({ context, onClose }: { context: ContextEntry; onClose: () => void }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState({ name: context.name, text: context.text });
  return (
    <Drawer
      title={context.name}
      wide
      onClose={onClose}
      footer={
        editing ? (
          <>
            <button type="button" className="btn-primary" onClick={() => void put("contexts", { ...context, ...draft, updatedAt: Date.now() }).then(onClose)}>
              Save
            </button>
            <button type="button" className="btn" onClick={() => setEditing(false)}>
              Cancel
            </button>
          </>
        ) : (
          <button type="button" className="btn" onClick={() => setEditing(true)}>
            Edit
          </button>
        )
      }
    >
      {editing ? (
        <>
          <input className="input" value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} />
          <textarea className="input min-h-[60vh] py-2 font-mono text-[13px]" value={draft.text} onChange={(e) => setDraft({ ...draft, text: e.target.value })} />
        </>
      ) : (
        <Markdown>{context.text}</Markdown>
      )}
    </Drawer>
  );
}

/* ----------------------------------- Add ----------------------------------- */

function AddDrawer({ tab, onClose }: { tab: Tab; onClose: () => void }) {
  const [values, setValues] = useState<Record<string, string>>({});
  const [mode, setMode] = useState<"link" | "excel">("link");
  const [state, setState] = useState<{ busy?: string; error?: string }>({});
  const set = (k: string) => (e: { target: { value: string } }) => setValues({ ...values, [k]: e.target.value });
  const v = (k: string) => values[k]?.trim() ?? "";
  const field = (k: string, label: string, placeholder: string, hint?: string) => (
    <label className="flex flex-col gap-1.5">
      <span className="label">{label}</span>
      <input className="input" placeholder={placeholder} value={values[k] ?? ""} onChange={set(k)} />
      {hint && <span className="text-[12px] text-muted-foreground">{hint}</span>}
    </label>
  );
  const submit = (busy: string, task: () => Promise<unknown>) => {
    setState({ busy });
    task().then(onClose, (e) => setState({ error: errorText(e) }));
  };
  const title = { reports: "Add a dashboard", datasets: "Add a dataset", schemas: "Add a schema", contexts: "Add a context" }[tab];
  let body: ReactNode, ready = false, action: () => void = () => undefined;
  if (tab === "reports") {
    body = field("link", "Power BI report link", "https://app.powerbi.com/groups/…/reports/…", "Its semantic model is added with it. Reports of Power BI apps are view only.");
    ready = !!v("link");
    action = () => submit("Reading the report…", () => addReport(v("link")));
  } else if (tab === "datasets") {
    body = (
      <>
        <div className="flex gap-1">
          <button type="button" className="tab" data-active={mode === "link"} onClick={() => setMode("link")}>Semantic model</button>
          <button type="button" className="tab" data-active={mode === "excel"} onClick={() => setMode("excel")}>From Excel files</button>
        </div>
        {mode === "link" ? (
          field("link", "Semantic model link", "https://app.powerbi.com/groups/…/datasets/…")
        ) : (
          <>
            {field("name", "Name of the new model", "Sales model")}
            <label className="flex flex-col gap-1.5">
              <span className="label">Excel files (one sharing link per line; "link | table" to choose the table)</span>
              <textarea className="input h-28 py-2" placeholder={"https://…sharepoint.com/:x:/…\nhttps://…sharepoint.com/:x:/… | Orders"} value={values.excels ?? ""} onChange={set("excels")} />
              <span className="text-[12px] text-muted-foreground">Created in My workspace, one model table per Excel table (typed after its columns). Its data comes with its first refresh.</span>
            </label>
          </>
        )}
      </>
    );
    ready = mode === "link" ? !!v("link") : !!v("name") && !!v("excels");
    action = () =>
      mode === "link"
        ? submit("Reading the model…", () => addDataset(v("link")))
        : submit("Creating the model…", () => createDatasetFromExcel(v("name"), v("excels").split("\n").filter((l) => l.trim()).map((l) => ({ link: l.split("|")[0].trim(), table: l.split("|")[1]?.trim() }))));
  } else if (tab === "schemas") {
    body = (
      <>
        {field("excel", "Excel file (required)", "Sharing link of an .xlsx", "Its columns are detected (types guessed from the data, or read from its _schema sheet).")}
        {field("table", "Table (optional)", "Name of the Excel table; the first one by default")}
        {field("form", "Microsoft Form (optional)", "Link of the form filling it", "Its questions give the types and the choices of the matching columns.")}
        {field("title", "Title (optional)", "Name shown in the app")}
      </>
    );
    ready = !!v("excel");
    action = () => submit("Reading the file…", () => addSchema({ excel: v("excel"), table: v("table") || undefined, form: v("form") || undefined, title: v("title") || undefined }));
  } else {
    body = (
      <>
        {field("title", "Title", "e.g. Company vocabulary")}
        <label className="flex min-h-0 flex-1 flex-col gap-1.5">
          <span className="label">Content (markdown)</span>
          <textarea className="input min-h-[50vh] flex-1 resize-none overflow-y-auto py-2 font-mono text-[13px]" placeholder={"# Vocabulary\n- CA = [Revenue]\n…"} value={values.text ?? ""} onChange={set("text")} />
          <span className="text-[12px] text-muted-foreground">Given to the AI with the dashboards, datasets and schemas it is linked to.</span>
        </label>
      </>
    );
    ready = !!v("title") && !!v("text");
    action = () => submit("Saving…", () => put("contexts", { id: crypto.randomUUID(), name: v("title"), text: values.text!, updatedAt: Date.now() }));
  }
  return (
    <Drawer
      title={title}
      wide={tab === "contexts"}
      onClose={onClose}
      footer={
        <>
          <button type="button" className="btn-primary" disabled={!ready || !!state.busy} onClick={action}>
            {state.busy ?? "Add"}
          </button>
          <button type="button" className="btn" onClick={onClose}>
            Cancel
          </button>
          {tab === "reports" && <a className="ml-auto flex items-center gap-1 text-[13px] text-muted-foreground hover:underline" href="https://app.powerbi.com" target="_blank" rel="noreferrer">Power BI <ExternalLinkIcon className="size-3.5" /></a>}
        </>
      }
    >
      {body}
      {state.error && <p className="text-[13px] break-words text-destructive">{state.error}</p>}
    </Drawer>
  );
}
