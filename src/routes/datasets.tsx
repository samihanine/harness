import { useCallback, useEffect, useMemo, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { getCoreRowModel, getFilteredRowModel, getSortedRowModel, useReactTable, type ColumnDef, type Row as TRow, type SortingState } from "@tanstack/react-table";
import { z } from "zod";
import {
  ArrowDownIcon,
  ArrowUpIcon,
  DatabaseIcon,
  ExternalLinkIcon,
  FilterIcon,
  PanelLeftIcon,
  PlusIcon,
  RefreshCwIcon,
  SearchIcon,
  SlidersHorizontalIcon,
  TableIcon,
  Trash2Icon,
  UploadIcon,
  XIcon,
} from "lucide-react";
import { Chat } from "@/components/chat";
import { FieldInput, OptionBadges, Thumbnail, type Uploader } from "@/components/field-input";
import { Popover } from "@/components/popover";
import { Empty, SidePanel } from "@/components/side-panel";
import { tool, type Agent } from "@/agent/loop";
import { daxTool } from "@/agent/pbi-tools";
import { addRows, deleteRows, imagesFolderLabel, readTable, saveFields, updateRows, uploadImage, type Table } from "@/lib/excel";
import { API_REFRESHES_PER_DAY, addDataset, apiRefreshesToday, modelPage, refreshDatasetText } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { rowsText } from "@/lib/pbi";
import { COLORS, FIELD_TYPES, TINTS, TYPE_LABELS, isHidden, labelOf, listOf, titleField, type Field, type Row } from "@/lib/schema";
import { useCollection, type DatasetEntry, type ExcelEntry } from "@/lib/store";
import { createModel, readModelFiles, tableTmdl, writeModelFiles } from "@/lib/tmdl";
import { refreshModel } from "@/server/ms";
import { useLast } from "@/lib/last";

export const Route = createFileRoute("/datasets")({
  validateSearch: (s: Record<string, unknown>) => ({ excel: (s.excel as string) || undefined }),
  component: Datasets,
});

/* ---------------------------------- State ---------------------------------- */

const NO_ROWS: Row[] = [];

type View = { layout: "table" | "detail"; search: string; filters: Record<string, string[]>; sort: SortingState; selected?: number };
const INITIAL: View = { layout: "table", search: "", filters: {}, sort: [] };

/** View of a file, remembered in this browser. */
function useView(id = "") {
  const key = `dataset-view:${id}`;
  const [view, setView] = useState<View>(INITIAL);
  useEffect(() => {
    try {
      setView({ ...INITIAL, ...JSON.parse(localStorage.getItem(key) ?? "{}") });
    } catch {
      setView(INITIAL);
    }
  }, [key]);
  const update = useCallback((patch: Partial<View>) => setView((v) => {
    const next = { ...v, ...patch };
    try {
      localStorage.setItem(key, JSON.stringify(next));
    } catch {}
    return next;
  }), [key]);
  return [view, update] as const;
}

/** The Excel table, with in-place updates after each write. */
function useTable(excel?: ExcelEntry) {
  const [table, setTable] = useState<Table>();
  const [error, setError] = useState<string>();
  const reload = useCallback(async () => {
    if (!excel) return setTable(undefined);
    try {
      setTable(await readTable(excel));
      setError(undefined);
    } catch (e) {
      setError(errorText(e));
    }
  }, [excel?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => void reload(), [reload]);
  /** Runs a write, then shows the result (or the error). */
  const write = useCallback(async (task: (t: Table) => Promise<unknown>, { refetch = false } = {}) => {
    if (!table) return;
    try {
      await task(table);
      setError(undefined);
      if (refetch) await reload();
      else setTable({ ...table, rows: [...table.rows] });
    } catch (e) {
      setError(errorText(e));
      throw e;
    }
  }, [table, reload]);
  return { table, error, reload, write };
}

/* ---------------------------------- Page ----------------------------------- */

function Datasets() {
  const { excel: excelId } = Route.useSearch();
  const navigate = useNavigate({ from: "/datasets" });
  useLast("datasets:excel", excelId, (excel) => navigate({ search: { excel }, replace: true }));
  const excels = useCollection("excels");
  const excel = excels.find((e) => e.id === excelId);
  const dataset = useCollection("datasets").find((d) => d.excelId === excelId);
  const { table, error, reload, write } = useTable(excel);
  const [view, setView] = useView(excelId);
  const [dialog, setDialog] = useState<"new" | "fields">();

  const columns = useMemo<ColumnDef<Row>[]>(
    () =>
      (table?.fields ?? []).filter((f) => !isHidden(f)).map((f) => ({
        id: f.name,
        accessorFn: (row) => (f.type === "option" ? listOf(row[f.name]).join(", ") : row[f.name]),
        filterFn: (row, id, values: string[]) => !values?.length || listOf(row.original[id]).some((v) => values.includes(v)),
        sortUndefined: "last",
      })),
    [table?.fields],
  );
  // Stable references: TanStack recomputes (and resets its state) when they change.
  const columnFilters = useMemo(() => Object.entries(view.filters).map(([id, value]) => ({ id, value })), [view.filters]);
  const grid = useReactTable({
    data: table?.rows ?? NO_ROWS,
    columns,
    state: { globalFilter: view.search, sorting: view.sort, columnFilters },
    onSortingChange: (u) => setView({ sort: typeof u === "function" ? u(view.sort) : u }),
    // No pagination: never reset it (each reset is a state change, hence a new render).
    autoResetAll: false,
    globalFilterFn: "includesString",
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
  });
  const rows = grid.getRowModel().rows;
  const change = (index: number, name: string, value: unknown) => void write((t) => updateRows(excel!, t, [{ index, values: { [name]: value } }])).catch(() => undefined);
  const upload: Uploader | undefined = excel && { save: (file) => uploadImage(excel, file), label: imagesFolderLabel(excel) };

  const agent = useMemo(() => (excel && table ? datasetAgent({ excel, table, dataset, rows: () => grid.getRowModel().rows, view, setView, write }) : undefined), [excel, table, dataset, view, grid, setView, write]);

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={[
          { id: "chat", label: "AI", content: agent ? <Chat scope={`dataset:${excelId}`} agent={agent} placeholder="Ask to filter, fix or add rows, change fields, add measures…" /> : <Empty>Choose an Excel file.</Empty> },
          { id: "model", label: "Semantic model", content: excel && table ? <ModelTab excel={excel} table={table} dataset={dataset} /> : <Empty>Choose an Excel file.</Empty> },
        ]}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex h-12 shrink-0 items-center gap-1.5 border-b bg-card px-3">
          <select className="input w-44" value={excelId ?? ""} onChange={(e) => navigate({ search: { excel: e.target.value } })}>
            <option value="" disabled>
              Choose an Excel file…
            </option>
            {excels.map((e) => (
              <option key={e.id} value={e.id}>
                {e.name}
              </option>
            ))}
          </select>
          {excel && table && (
            <>
              <span className="text-[12px] text-muted-foreground tabular-nums">{rows.length === table.rows.length ? rows.length : `${rows.length} / ${table.rows.length}`} rows</span>
              <span className="ml-auto" />
              <label className="flex h-8 items-center gap-1.5 rounded-md border bg-card px-2 focus-within:border-ring">
                <SearchIcon className="size-3.5 text-muted-foreground" />
                <input value={view.search} onChange={(e) => setView({ search: e.target.value })} placeholder="Search" className="w-32 bg-transparent outline-none" />
              </label>
              <Filters fields={table.fields} view={view} setView={setView} />
              <Segmented
                value={view.layout}
                onChange={(layout) => setView({ layout })}
                items={[
                  { value: "table", label: "Table", icon: <TableIcon /> },
                  { value: "detail", label: "Detail", icon: <PanelLeftIcon /> },
                ]}
              />
              <button type="button" className="btn" onClick={() => setDialog("fields")}>
                <SlidersHorizontalIcon /> Fields
              </button>
              <button type="button" className="btn-primary" onClick={() => setDialog("new")}>
                <PlusIcon /> New
              </button>
              <button type="button" className="icon-btn" title="Reload from Excel" onClick={() => void reload()}>
                <RefreshCwIcon />
              </button>
              <a className="icon-btn" title="Open in Excel" href={excel.link} target="_blank" rel="noreferrer">
                <ExternalLinkIcon />
              </a>
            </>
          )}
        </div>
        {error && <p className="border-b bg-red-500/5 px-4 py-1.5 text-[12px] text-destructive">{error}</p>}
        {!excel && <Empty>Add Excel files (with an Excel table) in the Library.</Empty>}
        {excel && !table && !error && <Empty>Loading…</Empty>}
        {excel && table && table.fields.length > 0 && (
          view.layout === "table" ? (
            <TableLayout table={table} rows={rows} grid={grid} upload={upload} onChange={change} onOpen={(index) => setView({ layout: "detail", selected: index })} onDelete={(index) => void write((t) => deleteRows(excel, [index]), { refetch: true })} />
          ) : (
            <DetailLayout table={table} rows={rows} selected={view.selected} onSelect={(selected) => setView({ selected })} upload={upload} onChange={change} />
          )
        )}
      </div>
      {excel && table && dialog === "new" && (
        <NewRow fields={table.fields} upload={upload} onClose={() => setDialog(undefined)} onCreate={async (values) => {
          await write((t) => addRows(excel, t, [values]), { refetch: true });
          setView({ layout: "detail", selected: table.rows.length });
        }} />
      )}
      {excel && table && dialog === "fields" && (
        <FieldsEditor table={table} onClose={() => setDialog(undefined)} onSave={(fields) => write((t) => saveFields(excel, t, fields), { refetch: true })} />
      )}
    </div>
  );
}

function Segmented<T extends string>({ value, onChange, items }: { value: T; onChange: (v: T) => void; items: { value: T; label: string; icon: ReactNode }[] }) {
  return (
    <div className="flex rounded-md bg-muted p-0.5">
      {items.map((item) => (
        <button key={item.value} type="button" title={item.label} onClick={() => onChange(item.value)} className={`flex size-7 items-center justify-center rounded text-muted-foreground [&_svg]:size-3.5 ${value === item.value ? "bg-card text-foreground shadow-sm" : ""}`}>
          {item.icon}
        </button>
      ))}
    </div>
  );
}

function Filters({ fields, view, setView }: { fields: Field[]; view: View; setView: (p: Partial<View>) => void }) {
  const options = fields.filter((f) => f.type === "option");
  const active = Object.values(view.filters).filter((v) => v.length).length;
  if (!options.length) return null;
  return (
    <Popover width={240} trigger={(open) => (
      <button type="button" className={`btn ${active ? "" : "text-muted-foreground"}`} onClick={open}>
        <FilterIcon /> {active ? `${active} filter${active > 1 ? "s" : ""}` : "Filter"}
      </button>
    )}>
      {() => (
        <div className="flex max-h-96 flex-col gap-3 overflow-y-auto p-1">
          {options.map((f) => {
            const selected = view.filters[f.name] ?? [];
            return (
              <div key={f.name} className="flex flex-col gap-1">
                <span className="label uppercase">{labelOf(f)}</span>
                {f.options?.map((o) => {
                  const on = selected.includes(o.value);
                  return (
                    <label key={o.value} className="flex items-center gap-2 py-0.5">
                      <input type="checkbox" className="accent-foreground" checked={on} onChange={() => setView({ filters: { ...view.filters, [f.name]: on ? selected.filter((v) => v !== o.value) : [...selected, o.value] } })} />
                      <span className={`inline-flex h-5 items-center rounded-full px-2 text-[12px] ${TINTS[o.color]}`}>{o.value}</span>
                    </label>
                  );
                })}
              </div>
            );
          })}
          {active > 0 && (
            <button type="button" className="btn h-7 w-fit" onClick={() => setView({ filters: {} })}>
              Clear filters
            </button>
          )}
        </div>
      )}
    </Popover>
  );
}

type LayoutProps = { table: Table; rows: TRow<Row>[]; upload?: Uploader; onChange: (index: number, name: string, value: unknown) => void };

function TableLayout({ table, rows, grid, upload, onChange, onOpen, onDelete }: LayoutProps & { grid: ReturnType<typeof useReactTable<Row>>; onOpen: (index: number) => void; onDelete: (index: number) => void }) {
  const width = (f: Field) => (f.type === "text" ? "min-w-72 max-w-96" : f.type === "string" ? "min-w-44" : f.type === "image" ? "w-16" : "min-w-28");
  return (
    <div className="min-h-0 flex-1 overflow-auto bg-card">
      <table className="w-max min-w-full border-separate border-spacing-0">
        <thead className="sticky top-0 z-10 bg-card/95 backdrop-blur">
          {grid.getHeaderGroups().map((group) => (
            <tr key={group.id}>
              {group.headers.map((header) => {
                const f = table.fields.find((x) => x.name === header.id)!;
                const sorted = header.column.getIsSorted();
                return (
                  <th key={header.id} className={`border-r border-b p-0 text-left font-medium ${width(f)}`}>
                    <button type="button" title={f.description} onClick={header.column.getToggleSortingHandler()} className="flex h-8 w-full items-center gap-1 px-2 text-[12px] text-muted-foreground hover:text-foreground">
                      {labelOf(f)}
                      {f.required && <span className="text-destructive/70">*</span>}
                      {sorted === "asc" ? <ArrowUpIcon className="size-3" /> : sorted === "desc" ? <ArrowDownIcon className="size-3" /> : null}
                    </button>
                  </th>
                );
              })}
              <th className="w-16 border-b" />
            </tr>
          ))}
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.index} className="group hover:bg-muted/40">
              {row.getVisibleCells().map((cell) => {
                const f = table.fields.find((x) => x.name === cell.column.id)!;
                return (
                  <td key={cell.id} className={`border-r border-b p-0 align-top ${width(f)}`}>
                    <FieldInput variant="cell" field={f} value={row.original[f.name]} upload={upload} onChange={(v) => onChange(row.index, f.name, v)} />
                  </td>
                );
              })}
              <td className="border-b px-1 align-middle">
                <div className="flex opacity-0 group-hover:opacity-100">
                  <button type="button" className="icon-btn" title="Open" onClick={() => onOpen(row.index)}>
                    <PanelLeftIcon />
                  </button>
                  <button type="button" className="icon-btn" title="Delete row" onClick={() => onDelete(row.index)}>
                    <Trash2Icon />
                  </button>
                </div>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {rows.length === 0 && <Empty>No rows match this view.</Empty>}
    </div>
  );
}

function DetailLayout({ table, rows, selected, onSelect, upload, onChange }: LayoutProps & { selected?: number; onSelect: (index: number) => void }) {
  const title = titleField(table.fields)!;
  const image = table.fields.find((f) => f.type === "image");
  const tags = table.fields.filter((f) => f.type === "option");
  const current = rows.find((r) => r.index === selected) ?? rows[0];
  return (
    <div className="flex min-h-0 flex-1">
      <ul className="flex w-80 shrink-0 flex-col gap-1.5 overflow-y-auto border-r bg-muted/40 p-2">
        {rows.map((r) => (
          <li key={r.index}>
            <button type="button" onClick={() => onSelect(r.index)} className={`flex w-full items-start gap-2 rounded-lg border bg-card p-2 text-left shadow-sm hover:border-ring ${r.index === current?.index ? "border-ring ring-2 ring-ring/20" : ""}`}>
              {image && <Thumbnail value={r.original[image.name]} className="size-10 shrink-0" />}
              <span className="flex min-w-0 flex-col gap-1">
                <span className="truncate font-medium">{String(r.original[title.name] ?? "") || <span className="font-normal text-muted-foreground">Untitled</span>}</span>
                <span className="flex flex-wrap gap-1">
                  {tags.map((f) => (
                    <OptionBadges key={f.name} field={f} value={r.original[f.name]} />
                  ))}
                </span>
              </span>
            </button>
          </li>
        ))}
        {rows.length === 0 && <Empty>No rows match this view.</Empty>}
      </ul>
      <div className="min-w-0 flex-1 overflow-y-auto bg-card">
        {current && (
          <div key={current.index} className="mx-auto flex max-w-3xl flex-col gap-4 px-8 py-6">
            {table.fields.filter((f) => !isHidden(f)).map((f) => (
              <FormField key={f.name} field={f}>
                <FieldInput variant="form" field={f} value={current.original[f.name]} upload={upload} onChange={(v) => onChange(current.index, f.name, v)} />
              </FormField>
            ))}
            <p className="text-[11px] text-muted-foreground/70">Row {current.index + 1} of the Excel table</p>
          </div>
        )}
      </div>
    </div>
  );
}

// A <div>, not a <label>: a label forwards clicks to its first button (e.g. the editor's Bold).
const FormField = ({ field, children }: { field: Field; children: ReactNode }) => (
  <div className="flex flex-col gap-1.5">
    <span className="text-[12px] font-medium">
      {labelOf(field)}
      {field.required && <span className="text-destructive/70"> *</span>}
    </span>
    {children}
    {field.description && <span className="text-[11px] text-muted-foreground">{field.description}</span>}
  </div>
);

function Dialog({ title, onClose, children, wide }: { title: string; onClose: () => void; children: ReactNode; wide?: boolean }) {
  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center bg-black/30 p-6" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className={`flex max-h-full w-full flex-col overflow-hidden rounded-xl border bg-card shadow-xl ${wide ? "max-w-4xl" : "max-w-xl"}`}>
        <div className="flex items-center border-b px-4 py-2.5">
          <h2 className="font-medium">{title}</h2>
          <button type="button" className="icon-btn ml-auto" onClick={onClose}>
            <XIcon />
          </button>
        </div>
        <div className="flex flex-col gap-4 overflow-y-auto p-4">{children}</div>
      </div>
    </div>
  );
}

function NewRow({ fields, upload, onClose, onCreate }: { fields: Field[]; upload?: Uploader; onClose: () => void; onCreate: (values: Row) => Promise<void> }) {
  const [values, setValues] = useState<Row>({});
  const [error, setError] = useState<string>();
  return (
    <Dialog title="New row" onClose={onClose}>
      {fields.filter((f) => !isHidden(f)).map((f) => (
        <FormField key={f.name} field={f}>
          <FieldInput variant="form" field={f} value={values[f.name]} upload={upload} onChange={(v) => setValues((s) => ({ ...s, [f.name]: v }))} />
        </FormField>
      ))}
      {error && <p className="text-[12px] text-destructive">{error}</p>}
      <div className="flex justify-end gap-2">
        <button type="button" className="btn" onClick={onClose}>
          Cancel
        </button>
        <button type="button" className="btn-primary" onClick={() => onCreate(values).then(onClose, (e) => setError(errorText(e)))}>
          Create
        </button>
      </div>
    </Dialog>
  );
}

/** Fields (stored in the workbook's hidden "_schema" sheet): label, type, options and colors, rules. */
function FieldsEditor({ table, onClose, onSave }: { table: Table; onClose: () => void; onSave: (fields: Field[]) => Promise<unknown> }) {
  const [fields, setFields] = useState<Field[]>(table.fields);
  const [error, setError] = useState<string>();
  const set = (i: number, patch: Partial<Field>) => setFields(fields.map((f, k) => (k === i ? { ...f, ...patch } : f)));
  return (
    <Dialog title="Fields" onClose={onClose} wide>
      {!table.hasSchema && <p className="rounded-md bg-amber-500/10 p-2 text-[12px] text-amber-800">Guessed from the data: saving writes them to a hidden "_schema" sheet of the Excel file.</p>}
      {fields.map((f, i) => (
        <div key={i} className="grid grid-cols-[1fr_1fr_1fr_auto] items-start gap-2 rounded-lg border p-3">
          <label className="flex flex-col gap-1">
            <span className="label">Column</span>
            <input className="input" value={f.name} disabled={table.header.includes(f.name)} onChange={(e) => set(i, { name: e.target.value })} />
          </label>
          <label className="flex flex-col gap-1">
            <span className="label">Label</span>
            <input className="input" value={f.label ?? ""} placeholder={labelOf(f)} onChange={(e) => set(i, { label: e.target.value || undefined })} />
          </label>
          <label className="flex flex-col gap-1">
            <span className="label">Type</span>
            <select className="input" value={f.type} onChange={(e) => set(i, { type: e.target.value as Field["type"] })}>
              {FIELD_TYPES.map((t) => (
                <option key={t} value={t}>
                  {TYPE_LABELS[t]}
                </option>
              ))}
            </select>
          </label>
          <div className="flex flex-col gap-1 pt-5">
            <label className="flex items-center gap-1.5 text-[12px]">
              <input type="checkbox" className="accent-foreground" checked={!!f.required} onChange={(e) => set(i, { required: e.target.checked || undefined })} /> Required
            </label>
            {f.type === "option" && (
              <label className="flex items-center gap-1.5 text-[12px]">
                <input type="checkbox" className="accent-foreground" checked={!!f.multiple} onChange={(e) => set(i, { multiple: e.target.checked || undefined })} /> Multiple
              </label>
            )}
          </div>
          <input className="input col-span-3" value={f.description ?? ""} placeholder="Description (shown under the field, and in the Power BI model)" onChange={(e) => set(i, { description: e.target.value || undefined })} />
          <span />
          {f.type === "option" && (
            <div className="col-span-4 flex flex-wrap items-center gap-1.5">
              {(f.options ?? []).map((o, k) => (
                <span key={k} className={`inline-flex h-7 items-center gap-1 rounded-full pr-1 pl-1 text-[12px] ${TINTS[o.color]}`}>
                  <select className="h-5 rounded-full bg-transparent text-[11px] outline-none" value={o.color} onChange={(e) => set(i, { options: f.options!.map((x, j) => (j === k ? { ...x, color: e.target.value as never } : x)) })}>
                    {COLORS.map((c) => (
                      <option key={c}>{c}</option>
                    ))}
                  </select>
                  <input className="w-24 bg-transparent outline-none" value={o.value} onChange={(e) => set(i, { options: f.options!.map((x, j) => (j === k ? { ...x, value: e.target.value } : x)) })} />
                  <button type="button" className="opacity-60 hover:opacity-100" onClick={() => set(i, { options: f.options!.filter((_, j) => j !== k) })}>
                    <XIcon className="size-3" />
                  </button>
                </span>
              ))}
              <button type="button" className="btn h-7" onClick={() => set(i, { options: [...(f.options ?? []), { value: "New", color: COLORS[(f.options?.length ?? 0) % COLORS.length] }] })}>
                <PlusIcon /> Option
              </button>
            </div>
          )}
        </div>
      ))}
      <button type="button" className="btn w-fit" onClick={() => setFields([...fields, { name: `field_${fields.length + 1}`, type: "string" }])}>
        <PlusIcon /> Add a field (new Excel column)
      </button>
      {error && <p className="text-[12px] text-destructive">{error}</p>}
      <div className="flex justify-end gap-2">
        <button type="button" className="btn" onClick={onClose}>
          Cancel
        </button>
        <button type="button" className="btn-primary" onClick={() => onSave(fields.map((f) => ({ ...f, options: f.type === "option" ? (f.options ?? []).filter((o) => o.value.trim()) : undefined }))).then(onClose, (e) => setError(errorText(e)))}>
          Save to the Excel file
        </button>
      </div>
    </Dialog>
  );
}

/* --------------------------------- Model tab -------------------------------- */

/** Generates the semantic model from the Excel table, or refreshes / updates it. */
function ModelTab({ excel, table, dataset }: { excel: ExcelEntry; table: Table; dataset?: DatasetEntry }) {
  const [state, setState] = useState<{ busy?: string; error?: string; done?: string }>({});
  const [name, setName] = useState(excel.name);
  const [used, setUsed] = useState<number>();
  const count = () => dataset && void apiRefreshesToday(dataset).then(setUsed, () => setUsed(undefined));
  useEffect(count, [dataset?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  const left = used === undefined ? undefined : API_REFRESHES_PER_DAY - used;
  const run = (label: string, task: () => Promise<string>) => {
    setState({ busy: label });
    task()
      .then((done) => setState({ done }), (e) => setState({ error: errorText(e) }))
      .finally(count);
  };
  const refresh = async (d: { id: string; groupId?: string }) => {
    const r = await refreshModel({ data: { datasetId: d.id, groupId: d.groupId } });
    if (r.status !== "Completed") throw new Error(`Refresh ${r.status}: ${r.error ?? ""}`);
  };
  return (
    <div className="flex flex-col gap-3 p-3">
      <ul className="flex flex-col gap-0.5 text-[12px]">
        {table.fields.map((f) => (
          <li key={f.name} className="flex justify-between gap-2">
            <span className="truncate">{labelOf(f)}</span>
            <span className="text-muted-foreground">{TYPE_LABELS[f.type]}</span>
          </li>
        ))}
      </ul>
      {dataset ? (
        <>
          <a className="flex items-center gap-2 font-medium underline" href={modelPage(dataset)} target="_blank" rel="noreferrer">
            <DatabaseIcon className="size-4" /> {dataset.name}
          </a>
          <button type="button" className="btn-primary w-fit" disabled={!!state.busy || left === 0} onClick={() => run("Refreshing…", () => refresh(dataset).then(() => "Data published to Power BI."))}>
            <UploadIcon /> Publish the data (refresh)
          </button>
          <button
            type="button"
            className="btn w-fit"
            disabled={!!state.busy || left === 0}
            onClick={() =>
              run("Updating the columns…", async () => {
                const files = await readModelFiles(dataset.id, dataset.groupId);
                const path = `definition/tables/${excel.table}.tmdl`;
                await writeModelFiles(dataset.id, files.map((f) => (f.path === path ? { ...f, text: tableTmdl(excel, table, f.text) } : f)), dataset.groupId);
                await refresh(dataset);
                await refreshDatasetText(dataset);
                return "Model updated from the fields (measures kept).";
              })
            }
          >
            <RefreshCwIcon /> Update the model from the fields
          </button>
          <p className={`text-[12px] ${left === 0 ? "text-destructive" : "text-muted-foreground"}`}>
            {used === undefined ? "Refreshes today: unknown." : `${used} / ${API_REFRESHES_PER_DAY} refreshes by the app today (UTC day, Power BI Pro).`}
            {left === 0 && " Limit reached:"}
          </p>
          <a className="btn w-fit" href={modelPage(dataset)} target="_blank" rel="noreferrer">
            <ExternalLinkIcon /> Refresh now in Power BI (not limited)
          </a>
        </>
      ) : (
        <>
          <label className="label">Semantic model name</label>
          <input className="input" value={name} onChange={(e) => setName(e.target.value)} />
          <button
            type="button"
            className="btn-primary w-fit"
            disabled={!!state.busy || !name}
            onClick={() =>
              run("Creating the model…", async () => {
                const id = await createModel(excel, table, name);
                await refresh({ id });
                await addDataset(id, undefined, { excelId: excel.id, context: excel.context });
                return "Model created in My workspace and added to the library.";
              })
            }
          >
            <DatabaseIcon /> Generate the semantic model
          </button>
        </>
      )}
      {state.busy && <p className="text-[12px] text-muted-foreground">{state.busy}</p>}
      {state.done && <p className="text-[12px] text-green-700">{state.done}</p>}
      {state.error && <p className="text-[12px] text-destructive">{state.error}</p>}
    </div>
  );
}

/* ----------------------------------- AI ------------------------------------ */

const cell = z.union([z.string(), z.number(), z.boolean(), z.array(z.string()), z.null()]);

function datasetAgent({ excel, table, dataset, rows, view, setView, write }: {
  excel: ExcelEntry;
  table: Table;
  dataset?: DatasetEntry;
  rows: () => TRow<Row>[];
  view: View;
  setView: (p: Partial<View>) => void;
  write: (task: (t: Table) => Promise<unknown>, o?: { refetch?: boolean }) => Promise<void>;
}): Agent {
  const tablePath = `definition/tables/${excel.table}.tmdl`;
  const fieldsText = table.fields
    .map((f) => `- ${f.name} (${TYPE_LABELS[f.type]}${f.required ? ", required" : ""}${f.multiple ? ", multiple" : ""})${f.label ? ` label "${f.label}"` : ""}${f.options?.length ? ` options: ${f.options.map((o) => `${o.value} [${o.color}]`).join(", ")}` : ""}${f.description ? ` — ${f.description}` : ""}`)
    .join("\n");
  return {
    instructions: `You manage the content of an Excel table on SharePoint ("${excel.name}"), shown as a table or as detail cards, and the Power BI semantic model generated from it. Rows are identified by "index" (their position in the Excel table). Text fields hold markdown; option fields only accept their listed values (multiple: a list); image fields hold image links; dates are YYYY-MM-DD. Data changes reach Power BI after refresh_model.`,
    context: async () =>
      [
        `Excel "${excel.name}" (table ${excel.table})${excel.context ? `\nNotes: ${excel.context}` : ""}`,
        `Fields:\n${fieldsText}`,
        `${table.rows.length} rows; the user's view (${view.layout}) shows ${rows().length}${view.search ? `, search "${view.search}"` : ""}${Object.entries(view.filters).filter(([, v]) => v.length).map(([k, v]) => `, ${k} in (${v.join(", ")})`).join("")}`,
        dataset
          ? `Semantic model "${dataset.name}" generated from it:\n${dataset.model ?? ""}\nRefreshes by the app today: ${await apiRefreshesToday(dataset).catch(() => "?")} / ${API_REFRESHES_PER_DAY}. Group the changes, then refresh once.`
          : "No semantic model generated yet (the user can generate it in the Semantic model tab).",
      ].join("\n\n"),
    tools: [
      tool({
        name: "read_rows",
        description: "Reads rows with their index: those of the user's view, or all of them.",
        args: z.object({ all: z.boolean().optional(), offset: z.number().optional(), limit: z.number().optional() }),
        readOnly: true,
        run: async ({ all, offset = 0, limit = 30 }) => {
          const list = all ? table.rows.map((r, index) => ({ index, ...r })) : rows().map((r) => ({ index: r.index, ...r.original }));
          return rowsText(list.slice(offset, offset + limit).map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, Array.isArray(v) ? v.join("; ") : v]))), limit);
        },
      }),
      tool({
        name: "update_rows",
        description: "Changes values of rows (only the given fields).",
        args: z.object({ rows: z.array(z.object({ index: z.number(), values: z.record(z.string(), cell) })) }),
        run: async ({ rows: changes }) => {
          await write((t) => updateRows(excel, t, changes));
          return `${changes.length} row(s) updated.`;
        },
      }),
      tool({
        name: "add_rows",
        description: "Adds rows at the end of the table.",
        args: z.object({ rows: z.array(z.record(z.string(), cell)) }),
        run: async ({ rows: added }) => {
          await write((t) => addRows(excel, t, added), { refetch: true });
          return `${added.length} row(s) added.`;
        },
      }),
      tool({
        name: "delete_rows",
        description: "Deletes rows by index.",
        args: z.object({ indexes: z.array(z.number()) }),
        run: async ({ indexes }) => {
          await write(() => deleteRows(excel, indexes), { refetch: true });
          return `${indexes.length} row(s) deleted.`;
        },
      }),
      tool({
        name: "set_view",
        description: "Changes what the user sees: layout, search, option filters (field → kept values), sort.",
        args: z.object({
          layout: z.enum(["table", "detail"]).optional(),
          search: z.string().optional(),
          filters: z.record(z.string(), z.array(z.string())).optional(),
          sort: z.object({ field: z.string(), desc: z.boolean().optional() }).nullable().optional(),
          open_row: z.number().optional().describe("Index of the row to show in the detail layout"),
        }),
        run: async ({ layout, search, filters, sort, open_row }) => {
          setView({
            ...(layout && { layout }),
            ...(search !== undefined && { search }),
            ...(filters && { filters }),
            ...(sort !== undefined && { sort: sort ? [{ id: sort.field, desc: !!sort.desc }] : [] }),
            ...(open_row !== undefined && { layout: "detail", selected: open_row }),
          });
          return "View updated.";
        },
      }),
      tool({
        name: "set_fields",
        description: "Replaces the fields (all of them, in order): type, label, description, required, multiple, options with colors. New names add Excel columns. Saved in the workbook.",
        args: z.object({
          fields: z.array(
            z.object({
              name: z.string(),
              type: z.enum(FIELD_TYPES),
              label: z.string().optional(),
              description: z.string().optional(),
              required: z.boolean().optional(),
              multiple: z.boolean().optional(),
              options: z.array(z.object({ value: z.string(), color: z.enum(COLORS) })).optional(),
            }),
          ),
        }),
        run: async ({ fields }) => {
          await write((t) => saveFields(excel, t, fields), { refetch: true });
          return "Fields saved.";
        },
      }),
      ...(dataset
        ? [
            daxTool(() => dataset),
            tool({
              name: "read_model",
              description: "Reads the TMDL of the model table (columns, measures, source).",
              args: z.object({}),
              readOnly: true,
              run: async () => (await readModelFiles(dataset.id, dataset.groupId)).find((f) => f.path === tablePath)?.text ?? "Table file not found.",
            }),
            tool({
              name: "write_model",
              description: "Replaces the TMDL of the model table (read it first, send it whole, tabs for indentation), then refreshes the model.",
              args: z.object({ tmdl: z.string() }),
              run: async ({ tmdl }) => {
                const files = await readModelFiles(dataset.id, dataset.groupId);
                await writeModelFiles(dataset.id, files.map((f) => (f.path === tablePath ? { ...f, text: tmdl } : f)), dataset.groupId);
                const r = await refreshModel({ data: { datasetId: dataset.id, groupId: dataset.groupId } });
                await refreshDatasetText(dataset);
                return `Model updated; refresh ${r.status}${r.error ? `: ${r.error}` : ""}.`;
              },
            }),
            tool({
              name: "refresh_model",
              description: `Publishes the Excel data to Power BI (model refresh, ${API_REFRESHES_PER_DAY} per day): only once all the requested changes are done.`,
              args: z.object({}),
              run: async () => {
                if ((await apiRefreshesToday(dataset)) >= API_REFRESHES_PER_DAY) return `Daily limit reached: tell the user to click "Refresh now" on ${modelPage(dataset)} (not limited).`;
                const r = await refreshModel({ data: { datasetId: dataset.id, groupId: dataset.groupId } });
                return `Refresh ${r.status}${r.error ? `: ${r.error}` : ""}.`;
              },
            }),
          ]
        : []),
    ],
  };
}

