import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
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
  KanbanIcon,
  BookmarkIcon,
  PlusIcon,
  RefreshCwIcon,
  SearchIcon,
  Settings2Icon,
  SlidersHorizontalIcon,
  TableIcon,
  Trash2Icon,
  UploadIcon,
  XIcon,
} from "lucide-react";
import { Chat } from "@/components/chat";
import { DATASET_TABS, visible } from "@/config";
import { Picker } from "@/components/picker";
import { FieldInput, OptionBadges, Thumbnail, type Uploader } from "@/components/field-input";
import { Popover } from "@/components/popover";
import { Empty, SidePanel } from "@/components/side-panel";
import { tool, type Agent } from "@/agent/loop";
import { daxTool } from "@/agent/pbi-tools";
import { modelTools } from "@/agent/model-tools";
import { addRows, deleteRows, imagesFolderLabel, readTable, saveFields, updateRows, uploadImage, withRows, withUpdates, withoutRows, type Table } from "@/lib/excel";
import { API_REFRESHES_PER_DAY, addDataset, addExcel, apiRefreshesToday, modelPage, refreshDatasetText } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { rowsText } from "@/lib/pbi";
import { COLORS, FIELD_TYPES, TINTS, TYPE_LABELS, isHidden, labelOf, listOf, titleField, type Field, type Row } from "@/lib/schema";
import { patch, useCollection, type DatasetEntry, type ExcelEntry } from "@/lib/store";
import { createModel, readModelFiles, tableTmdl, writeModelFiles } from "@/lib/tmdl";
import { refreshModel } from "@/server/ms";
import { useLast } from "@/lib/last";
import { useTable } from "@/lib/use-table";

export const Route = createFileRoute("/datasets")({
  validateSearch: (s: Record<string, unknown>) => ({ excel: (s.excel as string) || undefined }),
  component: Datasets,
});

/* ---------------------------------- State ---------------------------------- */

const NO_ROWS: Row[] = [];

/** Column condition (column menu): text contains, number / date range, yes / no. */
type Cond = { text?: string; min?: string; max?: string; bool?: "yes" | "no" };
type View = {
  layout: "table" | "detail" | "kanban";
  /** Option field whose values are the kanban columns. */
  kanbanBy?: string;
  search: string;
  /** Option fields → kept values (Filter button and column menus). */
  filters: Record<string, string[]>;
  /** Other fields → condition (column menus). */
  where: Record<string, Cond>;
  sort: SortingState;
  selected?: number;
};
const INITIAL: View = { layout: "table", search: "", filters: {}, where: {}, sort: [] };
const isActive = (c?: Cond) => !!c && Object.values(c).some((v) => v !== undefined && v !== "");

/** Row filter of a column: kept option values and / or condition. */
function matches(field: Field | undefined, value: unknown, { values, cond }: { values?: string[]; cond?: Cond }) {
  if (values?.length && !listOf(value).some((v) => values.includes(v))) return false;
  if (!cond) return true;
  if (cond.text && !listOf(value).join(" ").toLowerCase().includes(cond.text.toLowerCase())) return false;
  const numeric = field?.type === "number" || field?.type === "integer";
  const cmp = (v: string) => (numeric ? Number(v) : v);
  if (cond.min !== undefined && cond.min !== "" && !(value !== null && value !== undefined && (numeric ? Number(value) : String(value)) >= cmp(cond.min))) return false;
  if (cond.max !== undefined && cond.max !== "" && !(value !== null && value !== undefined && (numeric ? Number(value) : String(value)) <= cmp(cond.max))) return false;
  if (cond.bool === "yes" && value !== true) return false;
  if (cond.bool === "no" && value === true) return false;
  return true;
}

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

/* ---------------------------------- Page ----------------------------------- */

function Datasets() {
  const { excel: excelId } = Route.useSearch();
  const navigate = useNavigate({ from: "/datasets" });
  useLast("datasets:excel", excelId, (excel) => navigate({ search: { excel }, replace: true }));
  const excels = useCollection("excels");
  const excel = excels.find((e) => e.id === excelId);
  const dataset = useCollection("datasets").find((d) => d.excelId === excelId);
  const { table, error, reload, write } = useTable(excel);
  /** Column order = order of the fields in the "_schema" sheet (the Excel columns do not move). */
  const moveField = (from: string, to: string) => {
    if (!excel || !table || from === to) return;
    const next = [...table.fields];
    const [moved] = next.splice(next.findIndex((f) => f.name === from), 1);
    next.splice(next.findIndex((f) => f.name === to) + (table.fields.findIndex((f) => f.name === from) < table.fields.findIndex((f) => f.name === to) ? 1 : 0), 0, moved);
    void write((t) => saveFields(excel, t, next, { style: false }), { optimistic: (t) => ({ ...t, fields: next }) });
  };
  const [view, setView] = useView(excelId);
  const [dialog, setDialog] = useState<"new" | "fields" | { field: string }>();
  /** Rows checked by the user (Excel indexes): the AI then works on them only. */
  const [checked, setChecked] = useState<number[]>([]);
  useEffect(() => setChecked([]), [excelId, table?.rows.length]);
  const toggle = (index: number) => setChecked((c) => (c.includes(index) ? c.filter((i) => i !== index) : [...c, index]));
  const remove = (indexes: number[]) => void write(() => deleteRows(excel!, indexes), { optimistic: (t) => withoutRows(t, indexes) }).catch(() => undefined);

  const columns = useMemo<ColumnDef<Row>[]>(
    () =>
      (table?.fields ?? []).filter((f) => !isHidden(f)).map((f) => ({
        id: f.name,
        accessorFn: (row) => (f.type === "option" ? listOf(row[f.name]).join(", ") : row[f.name]),
        filterFn: (row, id, filter) => matches(f, row.original[id], filter),
        sortUndefined: "last",
      })),
    [table?.fields],
  );
  // Stable references: TanStack recomputes (and resets its state) when they change.
  const columnFilters = useMemo(
    () => [...new Set([...Object.keys(view.filters), ...Object.keys(view.where ?? {})])].map((id) => ({ id, value: { values: view.filters[id], cond: view.where?.[id] } })),
    [view.filters, view.where],
  );
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
  const change = (index: number, name: string, value: unknown) => {
    const changes = [{ index, values: { [name]: value } }];
    void write((t) => updateRows(excel!, t, changes), { optimistic: (t) => withUpdates(t, changes) }).catch(() => undefined);
  };
  const upload: Uploader | undefined = excel && { save: (file) => uploadImage(excel, file), label: imagesFolderLabel(excel) };

  const agent = useMemo(() => (excel && table ? datasetAgent({ excel, table, dataset, rows: () => grid.getRowModel().rows, checked, view, setView, write }) : undefined), [excel, table, dataset, view, grid, setView, write, checked]);

  return (
    <div className="flex h-full">
      <SidePanel
        tabs={visible([
          { id: "ai", label: "AI", content: agent ? <Chat scope={`dataset:${excelId}`} agent={agent} placeholder="Ask to filter, fix or add rows, change fields, add measures…" /> : <Empty>Choose an Excel file.</Empty> },
          { id: "model", label: "Semantic model", content: excel && table ? <ModelTab excel={excel} table={table} dataset={dataset} /> : <Empty>Choose an Excel file.</Empty> },
        ], DATASET_TABS)}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        <div className="flex h-12 shrink-0 items-center gap-1.5 overflow-x-auto border-b bg-card px-3">
          <Picker
            className="w-64"
            items={excels.map((e) => ({ id: e.id, label: e.name, hint: e.table }))}
            value={excelId}
            onChange={(excel) => navigate({ search: { excel } })}
            placeholder="Choose an Excel file…"
            add={{ label: "Add an Excel file", placeholder: "Sharing link of an .xlsx (SharePoint / OneDrive)", run: async (link) => (await addExcel(link)).id }}
          />
          {excel && table && (
            <>
              {checked.length > 0 && (
                <span className="flex shrink-0 items-center gap-1 rounded-md bg-blue-500/10 py-0.5 pr-0.5 pl-2 text-[12px] whitespace-nowrap text-blue-800">
                  {checked.length} selected
                  <button type="button" className="icon-btn" title="Delete the selected rows" onClick={() => remove(checked)}>
                    <Trash2Icon />
                  </button>
                  <button type="button" className="icon-btn" title="Clear the selection" onClick={() => setChecked([])}>
                    <XIcon />
                  </button>
                </span>
              )}
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
                  { value: "kanban", label: "Kanban", icon: <KanbanIcon /> },
                ]}
              />
              <SavedViews excel={excel} view={view} setView={setView} />
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
            <TableLayout onMoveField={moveField} table={table} rows={rows} grid={grid} upload={upload} onChange={change} view={view} setView={setView} onEditField={(field) => setDialog({ field })} checked={checked} onCheck={toggle} onCheckAll={(all) => setChecked(all ? rows.map((r) => r.index) : [])} onOpen={(index) => setView({ layout: "detail", selected: index })} onDelete={(index) => remove([index])} />
          ) : view.layout === "kanban" ? (
            <KanbanLayout table={table} rows={rows} view={view} setView={setView} onChange={change} onOpen={(index) => setView({ layout: "detail", selected: index })} />
          ) : (
            <DetailLayout table={table} rows={rows} selected={view.selected} onSelect={(selected) => setView({ selected })} upload={upload} onChange={change} checked={checked} onCheck={toggle} onDelete={(index) => remove([index])} />
          )
        )}
      </div>
      {excel && table && dialog === "new" && (
        <NewRow fields={table.fields} upload={upload} onClose={() => setDialog(undefined)} onCreate={async (values) => {
          await write((t) => addRows(excel, t, [values]), { refetch: true, optimistic: (t) => withRows(t, [values]) });
          setView({ layout: "detail", selected: table.rows.length });
        }} />
      )}
      {excel && table && (dialog === "fields" || typeof dialog === "object") && (
        <FieldsEditor table={table} only={typeof dialog === "object" ? dialog.field : undefined} onClose={() => setDialog(undefined)} onSave={(fields) => write((t) => saveFields(excel, t, fields), { refetch: true })} />
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
  const active = Object.values(view.filters).filter((v) => v.length).length + Object.values(view.where ?? {}).filter(isActive).length;
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
            <button type="button" className="btn h-7 w-fit" onClick={() => setView({ filters: {}, where: {} })}>
              Clear filters
            </button>
          )}
        </div>
      )}
    </Popover>
  );
}

type LayoutProps = {
  table: Table;
  rows: TRow<Row>[];
  upload?: Uploader;
  onChange: (index: number, name: string, value: unknown) => void;
  checked: number[];
  onCheck: (index: number) => void;
  onDelete: (index: number) => void;
};
const Check = ({ on, onChange, title }: { on: boolean; onChange: () => void; title: string }) => (
  <input type="checkbox" title={title} className="size-3.5 accent-foreground" checked={on} onChange={onChange} onClick={(e) => e.stopPropagation()} />
);

/** What a named view keeps (not the opened row). */
const viewToSave = ({ selected: _, ...v }: View) => v;

/** Named views of the file, kept in IndexedDB with the file: apply, save the current one, delete. */
function SavedViews({ excel, view, setView }: { excel: ExcelEntry; view: View; setView: (p: Partial<View>) => void }) {
  const [name, setName] = useState("");
  const views = excel.savedViews ?? [];
  const save = () => {
    const n = name.trim();
    if (!n) return;
    void patch("excels", excel.id, { savedViews: [...views.filter((v) => v.name !== n), { name: n, view: viewToSave(view) }] });
    setName("");
  };
  return (
    <Popover width={260} trigger={(open) => (
      <button type="button" className="btn" onClick={open}>
        <BookmarkIcon /> Views
      </button>
    )}>
      {(close) => (
        <div className="flex flex-col gap-2 p-1">
          {views.length === 0 && <p className="text-[12px] text-muted-foreground">No saved view yet.</p>}
          {views.map((v) => (
            <div key={v.name} className="flex items-center gap-1">
              <button type="button" className="flex-1 truncate rounded-md px-2 py-1 text-left hover:bg-muted" onClick={() => (setView({ ...INITIAL, ...(v.view as Partial<View>) }), close())}>
                {v.name}
                <span className="ml-1.5 text-[11px] text-muted-foreground">{String(v.view.layout ?? "")}</span>
              </button>
              <button type="button" className="icon-btn" title="Delete this view" onClick={() => void patch("excels", excel.id, { savedViews: views.filter((x) => x.name !== v.name) })}>
                <Trash2Icon />
              </button>
            </div>
          ))}
          <div className="flex gap-1.5 border-t pt-2">
            <input className="input h-7" placeholder="Name of the current view" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && save()} />
            <button type="button" className="btn-primary h-7" disabled={!name.trim()} onClick={save}>
              Save
            </button>
          </div>
          <p className="text-[11px] text-muted-foreground">Keeps the layout, search, filters, sort and kanban column. Same name: replaced.</p>
        </div>
      )}
    </Popover>
  );
}

/** Kanban: one column per value of an option field (+ empty); drag a card to change its value. */
function KanbanLayout({ table, rows, view, setView, onChange, onOpen }: {
  table: Table;
  rows: TRow<Row>[];
  view: View;
  setView: (p: Partial<View>) => void;
  onChange: (index: number, name: string, value: unknown) => void;
  onOpen: (index: number) => void;
}) {
  const options = table.fields.filter((f) => f.type === "option" && !isHidden(f));
  const by = options.find((f) => f.name === view.kanbanBy) ?? options.find((f) => !f.multiple) ?? options[0];
  const [over, setOver] = useState<string | null>(null);
  if (!by) return <Empty>The kanban needs an option field: add one in Fields.</Empty>;
  const title = titleField(table.fields)!;
  const image = table.fields.find((f) => f.type === "image");
  const tags = options.filter((f) => f !== by);
  const lanes = [...(by.options ?? []).map((o) => ({ key: o.value, label: o.value, color: o.color })), { key: "", label: "No value", color: "gray" as const }];
  const laneOf = (r: TRow<Row>) => listOf(r.original[by.name])[0] ?? "";
  const drop = (lane: string, index: number) => {
    const current = listOf(table.rows[index]?.[by.name]);
    if ((current[0] ?? "") === lane) return;
    // Multiple field: the dragged card's first value is replaced by the lane.
    const next = by.multiple ? (lane ? [lane, ...current.slice(1).filter((v) => v !== lane)] : current.slice(1)) : lane || null;
    onChange(index, by.name, next);
  };
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <div className="flex h-10 shrink-0 items-center gap-2 border-b bg-card px-3 text-[12px]">
        <span className="text-muted-foreground">Columns from</span>
        <select className="input h-7 w-48" value={by.name} onChange={(e) => setView({ kanbanBy: e.target.value })}>
          {options.map((f) => (
            <option key={f.name} value={f.name}>
              {labelOf(f)}
            </option>
          ))}
        </select>
      </div>
      <div className="flex min-h-0 flex-1 gap-3 overflow-x-auto bg-muted/40 p-3">
        {lanes.map((lane) => {
          const cards = rows.filter((r) => laneOf(r) === lane.key);
          if (lane.key === "" && cards.length === 0) return null;
          return (
            <section
              key={lane.key}
              onDragOver={(e) => (e.preventDefault(), setOver(lane.key))}
              onDragLeave={() => setOver((o) => (o === lane.key ? null : o))}
              onDrop={(e) => {
                setOver(null);
                drop(lane.key, Number(e.dataTransfer.getData("text/plain")));
              }}
              className={`flex w-72 shrink-0 flex-col rounded-xl border bg-muted/60 ${over === lane.key ? "ring-2 ring-ring/40" : ""}`}
            >
              <header className="flex items-center gap-2 px-3 py-2">
                <span className={`inline-flex h-5 items-center rounded-full px-2 text-[12px] ${TINTS[lane.color]}`}>{lane.label}</span>
                <span className="text-[12px] text-muted-foreground tabular-nums">{cards.length}</span>
              </header>
              <div className="flex min-h-16 flex-1 flex-col gap-2 overflow-y-auto px-2 pb-2">
                {cards.map((r) => (
                  <article
                    key={r.index}
                    draggable
                    onDragStart={(e) => e.dataTransfer.setData("text/plain", String(r.index))}
                    onClick={() => onOpen(r.index)}
                    className="flex cursor-grab flex-col gap-2 rounded-lg border bg-card p-2.5 shadow-sm hover:border-ring active:cursor-grabbing"
                  >
                    {image && r.original[image.name] ? <Thumbnail value={r.original[image.name]} className="h-28 w-full" /> : null}
                    <span className="font-medium">{String(r.original[title.name] ?? "") || <span className="font-normal text-muted-foreground">Untitled</span>}</span>
                    {tags.length > 0 && (
                      <span className="flex flex-wrap gap-1">
                        {tags.map((f) => (
                          <OptionBadges key={f.name} field={f} value={r.original[f.name]} />
                        ))}
                      </span>
                    )}
                  </article>
                ))}
              </div>
            </section>
          );
        })}
      </div>
    </div>
  );
}

/** Settings of one column: its filter (by type), its sort, its definition. */
function ColumnMenu({ field: f, view, setView, onEdit }: { field: Field; view: View; setView: (p: Partial<View>) => void; onEdit: () => void }) {
  const values = view.filters[f.name] ?? [];
  const cond = view.where?.[f.name] ?? {};
  const filtered = values.length > 0 || isActive(cond);
  const setCond = (patch: Cond) => setView({ where: { ...view.where, [f.name]: { ...cond, ...patch } } });
  const clear = () => {
    const { [f.name]: _a, ...filters } = view.filters;
    const { [f.name]: _b, ...where } = view.where ?? {};
    setView({ filters, where });
  };
  const sort = view.sort.find((s) => s.id === f.name);
  const setSort = (desc?: boolean) => setView({ sort: desc === undefined ? [] : [{ id: f.name, desc }] });
  const range = (type: "number" | "date") => (
    <div className="flex items-center gap-1.5">
      <input type={type} className="input h-7" placeholder="From" value={cond.min ?? ""} onChange={(e) => setCond({ min: e.target.value })} />
      <span className="text-muted-foreground">–</span>
      <input type={type} className="input h-7" placeholder="To" value={cond.max ?? ""} onChange={(e) => setCond({ max: e.target.value })} />
    </div>
  );
  return (
    <Popover width={260} trigger={(open) => (
      <button type="button" title="Filter, sort, edit the field" onClick={open} className={`relative inline-flex size-6 shrink-0 items-center justify-center rounded text-muted-foreground hover:bg-muted hover:text-foreground [&_svg]:size-3.5 ${filtered ? "text-blue-700" : "opacity-0 group-hover/th:opacity-100"}`}>
        <Settings2Icon />
        {filtered && <span className="absolute top-0.5 right-0.5 size-1.5 rounded-full bg-blue-600" />}
      </button>
    )}>
      {() => (
        <div className="flex flex-col gap-3 p-1">
          <section className="flex flex-col gap-1.5">
            <span className="label uppercase">Filter</span>
            {f.type === "option" &&
              f.options?.map((o) => {
                const on = values.includes(o.value);
                return (
                  <label key={o.value} className="flex items-center gap-2">
                    <input type="checkbox" className="accent-foreground" checked={on} onChange={() => setView({ filters: { ...view.filters, [f.name]: on ? values.filter((v) => v !== o.value) : [...values, o.value] } })} />
                    <span className={`inline-flex h-5 items-center rounded-full px-2 text-[12px] ${TINTS[o.color]}`}>{o.value}</span>
                  </label>
                );
              })}
            {(f.type === "string" || f.type === "text" || f.type === "image") && (
              <input autoFocus className="input h-7" placeholder="Contains…" value={cond.text ?? ""} onChange={(e) => setCond({ text: e.target.value })} />
            )}
            {(f.type === "number" || f.type === "integer") && range("number")}
            {f.type === "date" && range("date")}
            {f.type === "boolean" && (
              <div className="flex rounded-md bg-muted p-0.5 text-[12px]">
                {([[undefined, "All"], ["yes", "Yes"], ["no", "No"]] as const).map(([v, label]) => (
                  <button key={label} type="button" onClick={() => setCond({ bool: v })} className={`flex-1 rounded py-1 ${cond.bool === v ? "bg-card shadow-sm" : "text-muted-foreground"}`}>
                    {label}
                  </button>
                ))}
              </div>
            )}
            {filtered && (
              <button type="button" className="w-fit text-[12px] text-muted-foreground underline" onClick={clear}>
                Clear this filter
              </button>
            )}
          </section>
          <section className="flex flex-col gap-1.5 border-t pt-2">
            <span className="label uppercase">Sort</span>
            <div className="flex rounded-md bg-muted p-0.5 text-[12px]">
              {([[false, "A → Z"], [true, "Z → A"], [undefined, "None"]] as const).map(([desc, label]) => (
                <button key={label} type="button" onClick={() => setSort(desc)} className={`flex-1 rounded py-1 ${(sort ? sort.desc : undefined) === desc ? "bg-card shadow-sm" : "text-muted-foreground"}`}>
                  {label}
                </button>
              ))}
            </div>
          </section>
          <button type="button" className="btn h-7 border-t" onClick={onEdit}>
            <SlidersHorizontalIcon /> Edit the field ({TYPE_LABELS[f.type]})
          </button>
        </div>
      )}
    </Popover>
  );
}

function TableLayout({ table, rows, grid, upload, onChange, view, setView, onEditField, onMoveField, checked, onCheck, onCheckAll, onOpen, onDelete }: LayoutProps & {
  onMoveField: (from: string, to: string) => void;
  grid: ReturnType<typeof useReactTable<Row>>;
  view: View;
  setView: (p: Partial<View>) => void;
  onEditField: (name: string) => void;
  onCheckAll: (all: boolean) => void;
  onOpen: (index: number) => void;
}) {
  const all = rows.length > 0 && rows.every((r) => checked.includes(r.index));
  // Drag a column header onto another to move it there.
  const [dragged, setDragged] = useState<string>();
  const [over, setOver] = useState<string>();
  const width = (f: Field) => (f.type === "text" ? "min-w-72 max-w-96" : f.type === "string" ? "min-w-44" : f.type === "image" ? "w-16" : "min-w-28");
  return (
    <div className="min-h-0 flex-1 overflow-auto bg-card">
      <table className="w-max min-w-full border-separate border-spacing-0">
        <thead className="sticky top-0 z-10 bg-card/95 backdrop-blur">
          {grid.getHeaderGroups().map((group) => (
            <tr key={group.id}>
              <th className="w-9 border-r border-b">
                <span className="flex justify-center">
                  <Check on={all} onChange={() => onCheckAll(!all)} title="Select all the rows shown" />
                </span>
              </th>
              {group.headers.map((header) => {
                const f = table.fields.find((x) => x.name === header.id)!;
                const sorted = header.column.getIsSorted();
                return (
                  <th
                    key={header.id}
                    draggable
                    title="Drag to move the column"
                    onDragStart={(e) => (e.dataTransfer.setData("text/plain", f.name), (e.dataTransfer.effectAllowed = "move"), setDragged(f.name))}
                    onDragOver={(e) => dragged && (e.preventDefault(), setOver(f.name))}
                    onDragLeave={() => setOver((o) => (o === f.name ? undefined : o))}
                    onDrop={(e) => (e.preventDefault(), dragged && onMoveField(dragged, f.name), setDragged(undefined), setOver(undefined))}
                    onDragEnd={() => (setDragged(undefined), setOver(undefined))}
                    className={`group/th cursor-grab border-r border-b p-0 text-left font-medium active:cursor-grabbing ${width(f)} ${dragged === f.name ? "opacity-40" : ""} ${over === f.name && dragged !== f.name ? (table.fields.findIndex((x) => x.name === dragged) < table.fields.findIndex((x) => x.name === f.name) ? "shadow-[inset_-2px_0_0_var(--color-blue-600)]" : "shadow-[inset_2px_0_0_var(--color-blue-600)]") : ""}`}
                  >
                    <div className="flex h-8 items-center pr-1">
                      <button type="button" title={f.description} onClick={header.column.getToggleSortingHandler()} className="flex h-full min-w-0 flex-1 items-center gap-1 px-2 text-[12px] text-muted-foreground hover:text-foreground">
                        <span className="truncate">{labelOf(f)}</span>
                        {f.required && <span className="text-destructive/70">*</span>}
                        {sorted === "asc" ? <ArrowUpIcon className="size-3 shrink-0" /> : sorted === "desc" ? <ArrowDownIcon className="size-3 shrink-0" /> : null}
                      </button>
                      <ColumnMenu field={f} view={view} setView={setView} onEdit={() => onEditField(f.name)} />
                    </div>
                  </th>
                );
              })}
              <th className="w-16 border-b" />
            </tr>
          ))}
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.index} className={`group hover:bg-muted/40 ${checked.includes(row.index) ? "bg-blue-500/5" : ""}`}>
              <td className="border-r border-b align-middle">
                <span className="flex justify-center">
                  <Check on={checked.includes(row.index)} onChange={() => onCheck(row.index)} title="Select this row" />
                </span>
              </td>
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

function DetailLayout({ table, rows, selected, onSelect, upload, onChange, checked, onCheck, onDelete }: LayoutProps & { selected?: number; onSelect: (index: number) => void }) {
  const title = titleField(table.fields)!;
  const image = table.fields.find((f) => f.type === "image");
  const tags = table.fields.filter((f) => f.type === "option");
  const current = rows.find((r) => r.index === selected) ?? rows[0];
  return (
    <div className="flex min-h-0 flex-1">
      <ul className="flex w-80 shrink-0 flex-col gap-1.5 overflow-y-auto border-r bg-muted/40 p-2">
        {rows.map((r) => (
          <li key={r.index}>
            <div role="button" tabIndex={0} onClick={() => onSelect(r.index)} onKeyDown={(e) => e.key === "Enter" && onSelect(r.index)} className={`flex w-full cursor-pointer items-start gap-2 rounded-lg border p-2 text-left shadow-sm hover:border-ring ${checked.includes(r.index) ? "bg-blue-500/5" : "bg-card"} ${r.index === current?.index ? "border-ring ring-2 ring-ring/20" : ""}`}>
              <span className="pt-0.5">
                <Check on={checked.includes(r.index)} onChange={() => onCheck(r.index)} title="Select this row" />
              </span>
              {image && <Thumbnail value={r.original[image.name]} className="size-10 shrink-0" />}
              <span className="flex min-w-0 flex-col gap-1">
                <span className="truncate font-medium">{String(r.original[title.name] ?? "") || <span className="font-normal text-muted-foreground">Untitled</span>}</span>
                <span className="flex flex-wrap gap-1">
                  {tags.map((f) => (
                    <OptionBadges key={f.name} field={f} value={r.original[f.name]} />
                  ))}
                </span>
              </span>
            </div>
          </li>
        ))}
        {rows.length === 0 && <Empty>No rows match this view.</Empty>}
      </ul>
      <div className="min-w-0 flex-1 overflow-y-auto bg-card">
        {current && (
          <div key={current.index} className="mx-auto flex max-w-3xl flex-col gap-4 px-8 py-6">
            <div className="flex items-center gap-2">
              <h2 className="min-w-0 flex-1 truncate text-[16px] font-semibold">{String(current.original[title.name] ?? "") || "Untitled"}</h2>
              <button type="button" className="btn text-destructive" onClick={() => confirm("Delete this row from the Excel file?") && onDelete(current.index)}>
                <Trash2Icon /> Delete
              </button>
            </div>
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
const swap = <T,>(list: T[], a: number, b: number) => list.map((x, k) => (k === a ? list[b] : k === b ? list[a] : x));

function FieldsEditor({ table, only, onClose, onSave }: { table: Table; only?: string; onClose: () => void; onSave: (fields: Field[]) => Promise<unknown> }) {
  const [fields, setFields] = useState<Field[]>(table.fields);
  const [error, setError] = useState<string>();
  const set = (i: number, patch: Partial<Field>) => setFields(fields.map((f, k) => (k === i ? { ...f, ...patch } : f)));
  return (
    <Dialog title={only ? `Field: ${labelOf(table.fields.find((f) => f.name === only) ?? { name: only, type: "string" })}` : "Fields"} onClose={onClose} wide>
      {!table.hasSchema && <p className="rounded-md bg-amber-500/10 p-2 text-[12px] text-amber-800">Guessed from the data: saving writes them to a hidden "_schema" sheet of the Excel file.</p>}
      {fields.map((f, i) => (only && f.name !== only && table.fields[i]?.name !== only ? null :
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
          <div className="flex flex-col items-center">
            {!only && (
              <>
                <button type="button" className="icon-btn" title="Move up" disabled={i === 0} onClick={() => setFields(swap(fields, i, i - 1))}>
                  <ArrowUpIcon />
                </button>
                <button type="button" className="icon-btn" title="Move down" disabled={i === fields.length - 1} onClick={() => setFields(swap(fields, i, i + 1))}>
                  <ArrowDownIcon />
                </button>
              </>
            )}
            <button type="button" className="icon-btn" title="Remove the field (the Excel column is kept, only hidden)" onClick={() => setFields(fields.filter((_, k) => k !== i))}>
              <Trash2Icon />
            </button>
          </div>
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
      {!only && <button type="button" className="btn w-fit" onClick={() => setFields([...fields, { name: `field_${fields.length + 1}`, type: "string" }])}>
        <PlusIcon /> Add a field (new Excel column)
      </button>}
      <p className="text-[11px] text-muted-foreground">Removing a field never deletes its Excel column: it is only hidden here. Saving also restyles the Excel sheet (widths, dropdowns, option colors).</p>
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

function datasetAgent({ excel, table, dataset, rows, checked, view, setView, write }: {
  excel: ExcelEntry;
  table: Table;
  dataset?: DatasetEntry;
  rows: () => TRow<Row>[];
  checked: number[];
  view: View;
  setView: (p: Partial<View>) => void;
  write: (task: (t: Table) => Promise<unknown>, o?: { refetch?: boolean; optimistic?: (t: Table) => Table }) => Promise<unknown>;
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
        checked.length
          ? `THE USER SELECTED ${checked.length} ROW(S): indexes ${checked.join(", ")}. Act on these rows only (read them with read_rows), unless the user clearly asks otherwise.`
          : "No row selected: the request concerns all the rows.",
        `${table.rows.length} rows; the user's view (${view.layout}) shows ${rows().length}${view.search ? `, search "${view.search}"` : ""}${Object.entries(view.filters).filter(([, v]) => v.length).map(([k, v]) => `, ${k} in (${v.join(", ")})`).join("")}${Object.entries(view.where ?? {}).filter(([, c]) => isActive(c)).map(([k, c]) => `, ${k} ${JSON.stringify(c)}`).join("")}`,
        dataset
          ? `Semantic model "${dataset.name}" generated from it:\n${dataset.model ?? ""}\nRefreshes by the app today: ${await apiRefreshesToday(dataset).catch(() => "?")} / ${API_REFRESHES_PER_DAY}. Group the changes, then refresh once.`
          : "No semantic model generated yet (the user can generate it in the Semantic model tab).",
      ].join("\n\n"),
    tools: [
      tool({
        name: "read_rows",
        description: "Reads rows with their index: the selected rows when the user selected some, else those of the user's view; all=true for every row.",
        args: z.object({ all: z.boolean().optional(), offset: z.number().optional(), limit: z.number().optional() }),
        readOnly: true,
        run: async ({ all, offset = 0, limit = 30 }) => {
          // Only the fields (hidden Excel columns are not shown).
          const pick = (index: number, r: Row) => ({ index, ...Object.fromEntries(table.fields.map((f) => [f.name, r[f.name]])) });
          const list = all
            ? table.rows.map((r, index) => pick(index, r))
            : checked.length
              ? checked.map((index) => pick(index, table.rows[index]))
              : rows().map((r) => pick(r.index, r.original));
          return rowsText(list.slice(offset, offset + limit).map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, Array.isArray(v) ? v.join("; ") : v]))), limit);
        },
      }),
      tool({
        name: "update_rows",
        description: "Changes values of rows (only the given fields).",
        args: z.object({ rows: z.array(z.object({ index: z.number(), values: z.record(z.string(), cell) })) }),
        run: async ({ rows: changes }) => {
          await write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) });
          return `${changes.length} row(s) updated.`;
        },
      }),
      tool({
        name: "add_rows",
        description: "Adds rows at the end of the table.",
        args: z.object({ rows: z.array(z.record(z.string(), cell)) }),
        run: async ({ rows: added }) => {
          await write((t) => addRows(excel, t, added), { refetch: true, optimistic: (t) => withRows(t, added) });
          return `${added.length} row(s) added.`;
        },
      }),
      tool({
        name: "delete_rows",
        description: "Deletes rows by index.",
        args: z.object({ indexes: z.array(z.number()) }),
        run: async ({ indexes }) => {
          await write(() => deleteRows(excel, indexes), { optimistic: (t) => withoutRows(t, indexes) });
          return `${indexes.length} row(s) deleted.`;
        },
      }),
      tool({
        name: "set_view",
        description: "Changes what the user sees: layout, search, option filters (field → kept values), sort.",
        args: z.object({
          layout: z.enum(["table", "detail", "kanban"]).optional(),
          kanban_by: z.string().optional().describe("Option field used as kanban columns"),
          search: z.string().optional(),
          filters: z.record(z.string(), z.array(z.string())).optional().describe("Option fields → kept values"),
          where: z
            .record(z.string(), z.object({ text: z.string().optional(), min: z.string().optional(), max: z.string().optional(), bool: z.enum(["yes", "no"]).optional() }))
            .optional()
            .describe('Other fields → condition: {"text": "contains"} | {"min": "10", "max": "20"} (numbers or YYYY-MM-DD) | {"bool": "yes"}'),
          sort: z.object({ field: z.string(), desc: z.boolean().optional() }).nullable().optional(),
          open_row: z.number().optional().describe("Index of the row to show in the detail layout"),
        }),
        run: async ({ layout, kanban_by, search, filters, where, sort, open_row }) => {
          setView({
            ...(layout && { layout }),
            ...(kanban_by && { kanbanBy: kanban_by, layout: "kanban" }),
            ...(search !== undefined && { search }),
            ...(filters && { filters }),
            ...(where && { where }),
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
            ...modelTools(() => dataset),
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

