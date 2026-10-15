import { useEffect, useMemo, useState, type ReactNode } from "react";
import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { z } from "zod";
import {
  AlignLeftIcon,
  ArrowDownIcon,
  ArrowUpIcon,
  ArrowUpDownIcon,
  CalendarIcon,
  CheckIcon,
  ImageIcon,
  LinkIcon,
  ChevronDownIcon,
  ColumnsIcon,
  ExternalLinkIcon,
  FileIcon,
  HashIcon,
  KanbanIcon,
  ListIcon,
  PanelLeftCloseIcon,
  PanelLeftOpenIcon,
  PanelRightIcon,
  PlusIcon,
  RefreshCwIcon,
  SearchIcon,
  SlidersHorizontalIcon,
  SparklesIcon,
  StarIcon,
  Table2Icon,
  Trash2Icon,
  TypeIcon,
  XIcon,
} from "lucide-react";
import { Chat } from "@/components/chat";
import { Badge, FieldInput, FieldValue, type Uploader } from "@/components/field-input";
import { Picker } from "@/components/picker";
import { Popover } from "@/components/popover";
import { Empty, PageHeader, SidePanel } from "@/components/side-panel";
import { AI_PANEL, EXCEL_UPLOAD } from "@/config";
import { tool, type Agent } from "@/agent/loop";
import { withContexts } from "@/agent/context";
import { addRows, deleteRows, idError, updateRows, uploadFileTo, withRows, withUpdates, withoutRows, type Table } from "@/lib/excel";
import { rowsText } from "@/lib/pbi";
import { TYPE_LABELS, idPrefix, isIdField, labelOf, listOf, titleField, type Field, type FieldType, type Row } from "@/lib/schema";
import { patch, useCollection, type SchemaEntry } from "@/lib/store";
import { EMPTY_VIEW, NO_VALUE, OPS, OPS_OF, apply, filterLabel, matches, readView, shownFields, viewText, type Filter, type Layout, type TableView } from "@/lib/table-view";
import { last, useLast } from "@/lib/last";
import { useTable, type Write } from "@/lib/use-table";

export const Route = createFileRoute("/excel")({
  validateSearch: (s: Record<string, unknown>) => ({ schema: (s.schema as string) || undefined }),
  component: ExcelEditor,
});

const quiet = (p: Promise<unknown>) => void p.catch(() => undefined); // errors are shown by the page

export const TYPE_ICONS: Record<FieldType, ReactNode> = {
  string: <TypeIcon />,
  text: <AlignLeftIcon />,
  number: <HashIcon />,
  date: <CalendarIcon />,
  option: <ListIcon />,
  rating: <StarIcon />,
  url: <LinkIcon />,
  image: <ImageIcon />,
  file: <FileIcon />,
};

/* ---------------------------------- Views ---------------------------------- */

const ALL = "all";

/**
 * The views of a schema as pills: "All rows" (kept in this browser) and the named views (kept with the
 * schema in the library). The open view is changed in place; "+" saves it under a new name.
 */
function useViews(schema?: SchemaEntry) {
  const key = `table-view:${schema?.id}`;
  const [own, setOwn] = useState<TableView>(EMPTY_VIEW);
  const [active, setActiveState] = useState(ALL);
  useEffect(() => {
    try {
      setOwn(readView(JSON.parse(localStorage.getItem(key) ?? "{}")));
    } catch {
      setOwn(EMPTY_VIEW);
    }
    setActiveState(last.get(`${key}:active`) ?? ALL);
  }, [key]);
  const saved = schema?.views ?? [];
  const current = saved.find((v) => v.name === active);
  const view = current ? readView(current.view) : own;
  const setActive = (name: string) => {
    last.set(`${key}:active`, name);
    setActiveState(name);
  };
  const update = (change: Partial<TableView>) => {
    const next = { ...view, ...change };
    if (current && schema) return void patch("schemas", schema.id, { views: saved.map((v) => (v.name === current.name ? { ...v, view: next } : v)) });
    setOwn(next);
    try {
      localStorage.setItem(key, JSON.stringify(next));
    } catch {}
  };
  return {
    view,
    update,
    active: current ? current.name : ALL,
    names: saved.map((v) => v.name),
    open: setActive,
    saveAs: (name: string) => schema && void patch("schemas", schema.id, { views: [...saved.filter((v) => v.name !== name), { name, view }] }).then(() => setActive(name)),
    rename: (from: string, to: string) => schema && void patch("schemas", schema.id, { views: saved.map((v) => (v.name === from ? { ...v, name: to } : v)) }).then(() => active === from && setActive(to)),
    remove: (name: string) => schema && void patch("schemas", schema.id, { views: saved.filter((v) => v.name !== name) }).then(() => setActive(ALL)),
  };
}

/* ----------------------------------- Page ----------------------------------- */

/**
 * Excel editor: the content of a schema's Excel table (its structure is the Schema editor's), as a
 * table, a list with the open row's details, or a kanban; filters, sorts and named views. The AI
 * panel changes the rows and the view.
 */
function ExcelEditor() {
  const { schema: schemaId } = Route.useSearch();
  const navigate = useNavigate({ from: "/excel" });
  useLast("excel:schema", schemaId, (schema) => navigate({ search: { schema }, replace: true }));
  const schemas = useCollection("schemas").filter((s) => s.excel);
  const schema = schemas.find((s) => s.id === schemaId);
  const { table, error, reload, write } = useTable(schema);
  const views = useViews(schema);
  const { view } = views;
  const layout: Layout = view.layout ?? "table";
  /** The row open in the detail view (its index in the Excel table). */
  const [row, setRow] = useState<number>();
  /** Rows checked by the user (Excel indexes): the AI then works on them only. */
  const [checked, setChecked] = useState<number[]>([]);
  useEffect(() => (setChecked([]), setRow(undefined)), [schemaId]);

  const fields = table?.fields ?? [];
  const shown = useMemo(() => (table ? apply(table.rows, view, table.fields) : []), [table, view]);
  const excel = schema?.excel;
  /** A row shown with all its fields: the detail view, on that row. */
  const openRow = (index: number) => {
    setRow(index);
    if (layout !== "detail") views.update({ layout: "detail" });
  };
  const change = (index: number, name: string, value: unknown) => {
    const changes = [{ index, values: { [name]: value } }];
    quiet(write((t) => updateRows(excel!, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  const removeRows = (indexes: number[]) => {
    quiet(write(() => deleteRows(excel!, indexes), { optimistic: (t) => withoutRows(t, indexes) }));
    setChecked([]);
    setRow(undefined);
  };
  /** New: a row is added at once (next id, default values) and opened in the detail view to be filled. */
  const createRow = () => {
    if (!table) return;
    const index = table.rows.length;
    void write((t) => addRows(excel!, t, [{}], { partial: true }), { refetch: true, optimistic: (t) => withRows(t, [{}]) }).then(() => openRow(index), () => undefined);
    openRow(index);
  };
  // Uploads only when the schema has an uploads folder (set in the Schema editor) and the option is on.
  const folder = schema?.filesFolder;
  const upload: Uploader | undefined = EXCEL_UPLOAD && folder ? { save: (file) => uploadFileTo(folder, file), folder: folder.name } : undefined;
  const agent = useMemo(
    () => (schema && table ? excelAgent({ schema, table, shown: () => shown.map((s) => s.index), checked, view, setView: views.update, openRow, write }) : undefined),
    [schema, table, shown, checked, view, write], // eslint-disable-line react-hooks/exhaustive-deps
  );

  return (
    <div className="flex h-full">
      {AI_PANEL && (
        <SidePanel
          tabs={[{ id: "ai", label: "AI assistant", icon: SparklesIcon, content: agent ? <Chat scope={`excel:${schemaId}`} agent={agent} placeholder="Ask to find, fix or add rows, or to filter the view…" /> : <Empty>Choose a schema.</Empty> }]}
        />
      )}
      <div className="flex min-w-0 flex-1 flex-col">
        <PageHeader
          screen="Excel editor"
          picker={
            <Picker
              items={schemas.map((s) => ({ id: s.id, label: s.title, hint: s.builtIn ? "built-in" : undefined }))}
              value={schemaId}
              onChange={(schema) => navigate({ search: { schema } })}
              placeholder="Choose a schema…"
              empty={<>No schema with an Excel file: add one in the <Link to="/" search={{ tab: "schemas" }} className="underline">Library</Link>.</>}
            />
          }
          actions={
            schema && (
              <>
                {table && (
                  <button type="button" className="btn-primary h-9 rounded-full px-4" onClick={createRow}>
                    <PlusIcon /> New
                  </button>
                )}
                <button type="button" className="icon-btn" title="Read the file again" onClick={() => void reload()}>
                  <RefreshCwIcon />
                </button>
                <a className="btn h-9 rounded-full" href={excel!.webUrl ?? excel!.link} target="_blank" rel="noreferrer">
                  <ExternalLinkIcon /> Excel
                </a>
              </>
            )
          }
          below={schema && table && <FilterBar fields={fields} view={view} setView={views.update} count={shown.length} checked={checked} onClear={() => setChecked([])} onDelete={() => confirm(`Delete ${checked.length} row(s) from the Excel file?`) && removeRows(checked)} />}
        >
          {schema && table && <ViewTools fields={fields} views={views} />}
        </PageHeader>
        {error && <p className="border-b bg-red-500/5 px-6 py-2 text-[13px] text-destructive">{error}</p>}
        {!schema && <Empty>Choose a schema (an Excel table of the Library) to see and change its rows.</Empty>}
        {schema && !table && !error && <Empty>Loading…</Empty>}
        {schema && table && (
          <div className="min-h-0 flex-1 overflow-hidden p-5">
            {layout === "table" && <TableLayout table={table} items={shown} view={view} setView={views.update} upload={upload} onChange={change} checked={checked} setChecked={setChecked} />}
            {layout === "detail" && <DetailLayout key={schema.id} schema={schema} table={table} items={shown} selected={row} onSelect={setRow} write={write} upload={upload} onDelete={(i) => removeRows([i])} />}
            {layout === "kanban" && <KanbanLayout table={table} items={shown} view={view} onChange={change} onOpenRow={openRow} />}
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------- Toolbar ---------------------------------- */

const LAYOUTS: { id: Layout; label: string; icon: ReactNode }[] = [
  { id: "table", label: "Table", icon: <Table2Icon /> },
  { id: "detail", label: "List and details", icon: <PanelRightIcon /> },
  { id: "kanban", label: "Kanban", icon: <KanbanIcon /> },
];

/** First line: the views as pills (+ saves the current one), the layout, the columns shown. */
function ViewTools({ fields, views }: { fields: Field[]; views: ReturnType<typeof useViews> }) {
  const { view, update: setView } = views;
  const shown = shownFields(fields, view);
  return (
    <>
      {[ALL, ...views.names].map((name) => (
        <button
          key={name}
          type="button"
          className="tab group"
          data-active={name === views.active}
          onClick={() => views.open(name)}
          onDoubleClick={() => {
            if (name === ALL) return;
            const next = prompt("Rename the view", name)?.trim();
            if (next) views.rename(name, next);
          }}
        >
          {name === ALL ? "All rows" : name}
          {name !== ALL && name === views.active && <XIcon className="size-3.5 opacity-0 group-hover:opacity-70" onClick={(e) => (e.stopPropagation(), confirm(`Delete the view "${name}"?`) && views.remove(name))} />}
        </button>
      ))}
      <button type="button" className="icon-btn" title="Save the current filters, sorts and layout as a new view" onClick={() => { const name = prompt("Name of the new view")?.trim(); if (name) views.saveAs(name); }}>
        <PlusIcon />
      </button>
      <span className="mx-1 h-6 shrink-0 border-l" />
      <span className="flex shrink-0 rounded-full bg-muted p-0.5">
        {LAYOUTS.map((l) => (
          <button key={l.id} type="button" title={l.label} onClick={() => setView({ layout: l.id === "table" ? undefined : l.id })} className={`flex size-8 items-center justify-center rounded-full [&_svg]:size-4 ${(view.layout ?? "table") === l.id ? "bg-card text-foreground shadow-sm" : "text-muted-foreground hover:text-foreground"}`}>
            {l.icon}
          </button>
        ))}
      </span>
      {(view.layout ?? "table") === "table" && (
        <Popover width={240} trigger={(open) => (
          <button type="button" className="chip" onClick={open} title="Columns shown in this view">
            <ColumnsIcon /> {shown.length}/{fields.length}
          </button>
        )}>
          {() => (
            <div className="flex max-h-96 flex-col gap-1.5 overflow-y-auto p-1">
              {fields.map((f) => {
                const on = shown.includes(f);
                return (
                  <label key={f.name} className="flex items-center gap-2">
                    <input type="checkbox" className="size-4 accent-[var(--color-primary)]" checked={on} onChange={() => setView({ hidden: fields.filter((x) => (x === f ? on : !shown.includes(x))).map((x) => x.name) })} />
                    {labelOf(f)}
                  </label>
                );
              })}
              {view.hidden && (
                <button type="button" className="btn mt-1 h-8 w-fit" onClick={() => setView({ hidden: undefined })}>
                  Default (first 5)
                </button>
              )}
            </div>
          )}
        </Popover>
      )}
    </>
  );
}

/** Second line: filters (with their count), sorts, quick filters on the choices, search, results, selection. */
function FilterBar({ fields, view, setView, count, checked, onClear, onDelete }: { fields: Field[]; view: TableView; setView: (v: Partial<TableView>) => void; count: number; checked: number[]; onClear: () => void; onDelete: () => void }) {
  const choices = fields.filter((f) => f.type === "option").slice(0, 3);
  return (
    <>
      <Popover width={340} trigger={(open) => (
        <button type="button" className="chip" onClick={open}>
          <SlidersHorizontalIcon /> Filter
          {view.filters.length > 0 && <span className="flex size-5 items-center justify-center rounded-full bg-destructive text-[11.5px] text-white">{view.filters.length}</span>}
        </button>
      )}>
        {() => <FiltersEditor fields={fields} view={view} setView={setView} />}
      </Popover>
      <Popover width={280} trigger={(open) => (
        <button type="button" className="chip" onClick={open}>
          <ArrowUpDownIcon /> Sort
          {view.sorts.length > 0 && <span className="flex size-5 items-center justify-center rounded-full bg-foreground text-[11.5px] text-background">{view.sorts.length}</span>}
        </button>
      )}>
        {() => <SortsEditor fields={fields} view={view} setView={setView} />}
      </Popover>
      {view.layout === "kanban" && <GroupBy fields={fields} view={view} setView={setView} />}
      {choices.map((f) => {
        const current = view.filters.find((x) => x.field === f.name && x.op === "is");
        return (
          <Popover key={f.name} width={240} trigger={(open) => (
            <button type="button" className={`chip ${current?.values?.length ? "ring-2 ring-foreground/15" : ""}`} onClick={open}>
              {labelOf(f)}
              {current?.values?.length ? <span className="max-w-40 truncate text-muted-foreground">: {current.values.join(", ")}</span> : null}
              <ChevronDownIcon />
            </button>
          )}>
            {() => (
              <div className="flex flex-col">
                {f.options?.map((o) => {
                  const on = !!current?.values?.includes(o.value);
                  const values = on ? current!.values!.filter((v) => v !== o.value) : [...(current?.values ?? []), o.value];
                  const others = view.filters.filter((x) => x !== current);
                  return (
                    <button key={o.value} type="button" className="flex items-center gap-2 rounded-lg px-2 py-1.5 hover:bg-muted" onClick={() => setView({ filters: values.length ? [...others, { field: f.name, op: "is", values }] : others })}>
                      <input type="checkbox" readOnly checked={on} className="size-4 accent-[var(--color-primary)]" />
                      <Badge color={o.color}>{o.value}</Badge>
                    </button>
                  );
                })}
              </div>
            )}
          </Popover>
        );
      })}
      <span className="shrink-0 text-[13.5px] font-medium">
        {count} <span className="font-normal text-muted-foreground">results</span>
      </span>
      {checked.length > 0 && (
        <span className="flex h-9 items-center gap-1 rounded-full bg-foreground pr-1 pl-3.5 text-[13px] text-background">
          {checked.length} selected
          <button type="button" className="flex size-7 items-center justify-center rounded-full hover:bg-white/15" title="Delete the selected rows" onClick={onDelete}>
            <Trash2Icon className="size-4" />
          </button>
          <button type="button" className="flex size-7 items-center justify-center rounded-full hover:bg-white/15" title="Clear the selection" onClick={onClear}>
            <XIcon className="size-4" />
          </button>
        </span>
      )}
      <span className="ml-auto" />
      <label className="chip w-56 cursor-text">
        <SearchIcon />
        <input className="min-w-0 flex-1 bg-transparent outline-none placeholder:text-muted-foreground" placeholder="Search" value={view.search ?? ""} onChange={(e) => setView({ search: e.target.value })} />
      </label>
    </>
  );
}

/** Kanban: the choice field giving the columns (saved with the view). */
function GroupBy({ fields, view, setView }: { fields: Field[]; view: TableView; setView: (v: Partial<TableView>) => void }) {
  const choices = fields.filter((f) => f.type === "option" && !f.multiple);
  const current = kanbanField(fields, view);
  if (!current) return null;
  return (
    <Popover width={220} trigger={(open) => (
      <button type="button" className="chip" onClick={open} title="The choice field giving the kanban's columns">
        <ColumnsIcon /> Group by <span className="font-semibold">{labelOf(current)}</span>
        <ChevronDownIcon />
      </button>
    )}>
      {(close) => (
        <div className="flex flex-col">
          {choices.map((f) => (
            <button key={f.name} type="button" className="flex items-center gap-2 rounded-lg px-2 py-1.5 text-left hover:bg-muted" onClick={() => (setView({ group: f.name }), close())}>
              {labelOf(f)}
              {f === current && <CheckIcon className="ml-auto size-4" />}
            </button>
          ))}
        </div>
      )}
    </Popover>
  );
}

/** Every filter of the view: field, operator suited to its type, value(s). */
function FiltersEditor({ fields, view, setView }: { fields: Field[]; view: TableView; setView: (v: Partial<TableView>) => void }) {
  const set = (i: number, f?: Filter) => setView({ filters: f ? view.filters.map((x, k) => (k === i ? f : x)) : view.filters.filter((_, k) => k !== i) });
  return (
    <div className="flex flex-col gap-3 p-1">
      {view.filters.map((filter, i) => {
        const field = fields.find((f) => f.name === filter.field) ?? fields[0];
        return (
          <div key={i} className="flex flex-col gap-1.5 rounded-lg border p-2">
            <div className="flex gap-1.5">
              <select className="input h-8" value={field.name} onChange={(e) => {
                const next = fields.find((f) => f.name === e.target.value)!;
                set(i, { field: next.name, op: OPS_OF[next.type][0] });
              }}>
                {fields.map((f) => (
                  <option key={f.name} value={f.name}>{labelOf(f)}</option>
                ))}
              </select>
              <select className="input h-8 w-32" value={filter.op} onChange={(e) => set(i, { ...filter, op: e.target.value as Filter["op"] })}>
                {OPS_OF[field.type].map((o) => (
                  <option key={o} value={o}>{OPS[o]}</option>
                ))}
              </select>
              <button type="button" className="icon-btn" title="Remove the filter" onClick={() => set(i)}>
                <XIcon />
              </button>
            </div>
            {!NO_VALUE.includes(filter.op) &&
              (field.type === "option" ? (
                <div className="flex flex-wrap gap-1">
                  {field.options?.map((o) => {
                    const on = filter.values?.includes(o.value);
                    return (
                      <button key={o.value} type="button" className={on ? "" : "opacity-45 hover:opacity-100"} onClick={() => set(i, { ...filter, values: on ? filter.values!.filter((v) => v !== o.value) : [...(filter.values ?? []), o.value] })}>
                        <Badge color={o.color}>{o.value}</Badge>
                      </button>
                    );
                  })}
                </div>
              ) : (
                <input className="input h-8" type={field.type === "number" || field.type === "rating" ? "number" : field.type === "date" ? "date" : "text"} placeholder="Value" value={filter.value ?? ""} onChange={(e) => set(i, { ...filter, value: e.target.value })} />
              ))}
          </div>
        );
      })}
      {!view.filters.length && <p className="text-[13px] text-muted-foreground">No filter.</p>}
      <select className="input h-8" value="" onChange={(e) => {
        const f = fields.find((x) => x.name === e.target.value);
        if (f) setView({ filters: [...view.filters, { field: f.name, op: OPS_OF[f.type][0] }] });
      }}>
        <option value="">+ Add a filter on…</option>
        {fields.map((f) => (
          <option key={f.name} value={f.name}>{labelOf(f)}</option>
        ))}
      </select>
      {view.filters.length > 0 && <p className="text-[12px] text-muted-foreground">{view.filters.map((f) => filterLabel(f, fields)).join(" · ")}</p>}
    </div>
  );
}

function SortsEditor({ fields, view, setView }: { fields: Field[]; view: TableView; setView: (v: Partial<TableView>) => void }) {
  return (
    <div className="flex flex-col gap-2 p-1">
      {view.sorts.map((s, i) => {
        const field = fields.find((f) => f.name === s.field);
        return (
          <div key={i} className="flex items-center gap-1.5">
            <span className="flex-1 truncate">{field ? labelOf(field) : s.field}</span>
            <button type="button" className="btn h-8" onClick={() => setView({ sorts: view.sorts.map((x, k) => (k === i ? { ...x, desc: !x.desc } : x)) })}>
              {s.desc ? <ArrowDownIcon /> : <ArrowUpIcon />} {s.desc ? "Z → A" : "A → Z"}
            </button>
            <button type="button" className="icon-btn" onClick={() => setView({ sorts: view.sorts.filter((_, k) => k !== i) })}>
              <XIcon />
            </button>
          </div>
        );
      })}
      <select className="input h-8" value="" onChange={(e) => e.target.value && setView({ sorts: [...view.sorts, { field: e.target.value }] })}>
        <option value="">+ Sort by…</option>
        {fields.filter((f) => !view.sorts.some((s) => s.field === f.name)).map((f) => (
          <option key={f.name} value={f.name}>{labelOf(f)}</option>
        ))}
      </select>
    </div>
  );
}

/* ----------------------------------- Table ----------------------------------- */

function TableLayout({ table, items, view, setView, upload, onChange, checked, setChecked }: {
  table: Table;
  items: { row: Row; index: number }[];
  view: TableView;
  setView: (v: Partial<TableView>) => void;
  upload?: Uploader;
  onChange: (index: number, name: string, value: unknown) => void;
  checked: number[];
  setChecked: (c: number[]) => void;
}) {
  const columns = shownFields(table.fields, view);
  const all = items.length > 0 && items.every((i) => checked.includes(i.index));
  // Drag a header's right edge to resize the column (saved with the view when released).
  const [resizing, setResizing] = useState<{ name: string; width: number }>();
  const widthOf = (f: Field) => (resizing?.name === f.name ? resizing.width : view.widths?.[f.name]) ?? (isIdField(f) ? 130 : f.type === "text" ? 320 : f.type === "string" ? 220 : f.type === "file" || f.type === "url" ? 200 : f.type === "image" ? 110 : 170);
  const style = (f: Field) => ({ width: widthOf(f), minWidth: widthOf(f), maxWidth: widthOf(f) });
  const resize = (f: Field, e: React.PointerEvent<HTMLElement>) => {
    e.preventDefault();
    e.stopPropagation();
    const start = e.clientX;
    const from = widthOf(f);
    let current = from;
    const move = (m: PointerEvent) => setResizing({ name: f.name, width: (current = Math.max(70, Math.round(from + m.clientX - start))) });
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setResizing(undefined);
      setView({ widths: { ...view.widths, [f.name]: current } });
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };
  const sortBy = (f: Field) => {
    const s = view.sorts[0]?.field === f.name ? view.sorts[0] : undefined;
    setView({ sorts: s ? (s.desc ? [] : [{ field: f.name, desc: true }]) : [{ field: f.name }] });
  };
  return (
    <div className="card h-full overflow-auto">
      <table className="w-max min-w-full border-separate border-spacing-0">
        <thead className="sticky top-0 z-10 bg-card">
          <tr>
            <th className="w-12 border-r border-b bg-muted/40">
              <span className="flex justify-center">
                <input type="checkbox" className="size-4 accent-[var(--color-primary)]" title="Select all the rows shown" checked={all} onChange={() => setChecked(all ? [] : items.map((i) => i.index))} />
              </span>
            </th>
            {columns.map((f) => {
              const sort = view.sorts.find((s) => s.field === f.name);
              return (
                <th key={f.name} style={style(f)} className="relative border-r border-b bg-muted/40 p-0 last:border-r-0">
                  <button type="button" title={f.description ?? "Sort"} onClick={() => sortBy(f)} className="th flex w-full items-center gap-2 hover:bg-muted [&_svg]:size-4 [&_svg]:shrink-0 [&_svg]:text-muted-foreground">
                    {TYPE_ICONS[f.type]}
                    <span className="truncate">{labelOf(f)}</span>
                    {f.required && <span className="text-destructive/70">*</span>}
                    {sort && (sort.desc ? <ArrowDownIcon /> : <ArrowUpIcon />)}
                  </button>
                  <div role="separator" aria-orientation="vertical" title="Drag to resize" onPointerDown={(e) => resize(f, e)} className={`absolute inset-y-0 -right-1 z-10 w-2 cursor-col-resize hover:bg-zinc-400/40 ${resizing?.name === f.name ? "bg-zinc-400/50" : ""}`} />
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {items.map(({ row, index }) => (
            <tr key={index} className={`group ${checked.includes(index) ? "bg-muted/60" : "hover:bg-muted/40"}`}>
              <td className="border-r border-b align-middle">
                <span className="flex justify-center">
                  <input type="checkbox" className="size-4 accent-[var(--color-primary)]" title="Select this row" checked={checked.includes(index)} onChange={() => setChecked(checked.includes(index) ? checked.filter((i) => i !== index) : [...checked, index])} />
                </span>
              </td>
              {columns.map((f) => (
                <td key={f.name} style={style(f)} className="overflow-hidden border-r border-b p-0 align-middle last:border-r-0">
                  <FieldInput variant="cell" field={f} value={row[f.name]} upload={upload} validate={(v) => idError(table, index, v)} onChange={(v) => onChange(index, f.name, v)} />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {items.length === 0 && <p className="p-10 text-center text-muted-foreground">No row matches this view.</p>}
    </div>
  );
}

/* ---------------------------------- Detail ---------------------------------- */

/** The rows as a list (title and choices) and the open row's fields beside it. */
function DetailLayout({ schema, table, items, selected, onSelect, write, upload, onDelete }: { schema: SchemaEntry; table: Table; items: { row: Row; index: number }[]; selected?: number; onSelect: (index: number) => void; write: Write; upload?: Uploader; onDelete: (i: number) => void }) {
  // The list can be folded (remembered in this browser) to give the row all the width.
  const [folded, setFolded] = useState(() => last.get("detail-list") === "folded");
  const fold = (next: boolean) => (last.set("detail-list", next ? "folded" : "open"), setFolded(next));
  const title = titleField(table.fields);
  const id = table.fields.find(isIdField);
  const badges = table.fields.filter((f) => f.type === "option").slice(0, 2);
  useEffect(() => {
    if (selected === undefined && items[0]) onSelect(items[0].index);
  }, [items.length]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div className="flex h-full gap-4">
      {folded ? (
        <button type="button" className="card flex w-10 shrink-0 items-start justify-center pt-3 text-muted-foreground hover:text-foreground" title="Show the list" onClick={() => fold(false)}>
          <PanelLeftOpenIcon className="size-4" />
        </button>
      ) : (
      <div className="card flex w-96 shrink-0 flex-col overflow-hidden">
        <div className="flex h-11 shrink-0 items-center border-b px-3 text-[13px] text-muted-foreground">
          {items.length} rows
          <button type="button" className="icon-btn ml-auto size-7" title="Fold the list" onClick={() => fold(true)}>
            <PanelLeftCloseIcon />
          </button>
        </div>
        <div className="min-h-0 flex-1 overflow-y-auto">
        {items.map(({ row, index }) => (
          <button key={index} type="button" onClick={() => onSelect(index)} className={`flex w-full flex-col gap-1.5 border-b px-4 py-3 text-left last:border-b-0 ${selected === index ? "bg-accent" : "hover:bg-muted/40"}`}>
            <span className="flex items-baseline gap-2">
              {id && <span className="max-w-24 shrink-0 truncate text-[12.5px] text-muted-foreground tabular-nums">{String(row[id.name] ?? "")}</span>}
              <span className="font-medium break-words">{title ? String(row[title.name] ?? "") || "Untitled" : `Row ${index + 1}`}</span>
            </span>
            {badges.length > 0 && (
              <span className="flex flex-wrap gap-1">
                {badges.map((f) => listOf(row[f.name]).map((v) => <Badge key={`${f.name}${v}`} small color={f.options?.find((o) => o.value === v)?.color}>{v}</Badge>))}
              </span>
            )}
          </button>
        ))}
        {items.length === 0 && <p className="p-6 text-center text-muted-foreground">No row matches this view.</p>}
        </div>
      </div>
      )}
      <div className="card min-w-0 flex-1 overflow-y-auto p-6">
        {selected === undefined ? <Empty>Choose a row.</Empty> : <RowForm key={String(selected)} schema={schema} table={table} selected={selected} write={write} upload={upload} onDelete={onDelete} columns />}
      </div>
    </div>
  );
}

/* ---------------------------------- Kanban ---------------------------------- */

/** The kanban's columns: the view's grouping field, else the first single choice field. */
const kanbanField = (fields: Field[], view: TableView) => {
  const choices = fields.filter((f) => f.type === "option" && !f.multiple);
  return choices.find((f) => f.name === view.group) ?? choices[0];
};

/** One column per choice of a field (the view's grouping, saved with the view); drag a card to change its value. */
function KanbanLayout({ table, items, view, onChange, onOpenRow }: { table: Table; items: { row: Row; index: number }[]; view: TableView; onChange: (index: number, name: string, value: unknown) => void; onOpenRow: (index: number) => void }) {
  const field = kanbanField(table.fields, view);
  const [dragged, setDragged] = useState<number>();
  const [over, setOver] = useState<string>();
  if (!field) return <Empty>The kanban needs a single choice field (e.g. a status).</Empty>;
  const title = titleField(table.fields);
  const id = table.fields.find(isIdField);
  const others = table.fields.filter((f) => f !== field && f !== title && !isIdField(f) && ["option", "date", "rating"].includes(f.type)).slice(0, 3);
  const columns = [...(field.options ?? []).map((o) => ({ key: o.value, color: o.color })), { key: "", color: undefined }];
  return (
    <div className="flex h-full flex-col gap-3">
      <div className="flex min-h-0 flex-1 gap-4 overflow-x-auto pb-2">
        {columns.map((c) => {
          const cards = items.filter(({ row }) => (listOf(row[field.name])[0] ?? "") === c.key);
          if (!c.key && !cards.length) return null;
          return (
            <div
              key={c.key}
              onDragOver={(e) => (e.preventDefault(), setOver(c.key))}
              onDragLeave={() => setOver(undefined)}
              onDrop={(e) => {
                e.preventDefault();
                setOver(undefined);
                if (dragged !== undefined) onChange(dragged, field.name, c.key || null);
                setDragged(undefined);
              }}
              className={`flex w-80 shrink-0 flex-col gap-3 rounded-2xl p-3 transition-colors ${over === c.key ? "bg-zinc-200/70" : "bg-muted/60"}`}
            >
              <span className="flex items-center gap-2 px-1">
                {c.key ? <Badge color={c.color}>{c.key}</Badge> : <span className="font-medium text-muted-foreground">No value</span>}
                <span className="text-muted-foreground">{cards.length}</span>
              </span>
              <div className="flex min-h-0 flex-col gap-2.5 overflow-y-auto">
                {cards.map(({ row, index }) => (
                  <button
                    key={index}
                    type="button"
                    draggable
                    onDragStart={() => setDragged(index)}
                    onDragEnd={() => setDragged(undefined)}
                    onClick={() => onOpenRow(index)}
                    className={`flex flex-col gap-2 rounded-xl border bg-card p-3.5 text-left shadow-xs hover:shadow-md ${dragged === index ? "opacity-40" : ""}`}
                  >
                    {id && <span className="truncate text-[12px] text-muted-foreground tabular-nums">{String(row[id.name] ?? "")}</span>}
                    <span className="font-medium">{title ? String(row[title.name] ?? "") || "Untitled" : `Row ${index + 1}`}</span>
                    {others.length > 0 && (
                      <span className="flex flex-wrap items-center gap-1.5 text-[13px]">
                        {others.map((f) => (row[f.name] === null || row[f.name] === undefined || row[f.name] === "" ? null : <FieldValue key={f.name} field={f} value={row[f.name]} />))}
                      </span>
                    )}
                  </button>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* --------------------------------- Row form --------------------------------- */

/**
 * Every field of a row, as a form: long texts, short texts and files on the whole width, the other
 * fields side by side (`columns`: three per line). Changes are written at once.
 */
function RowForm({ schema, table, selected, write, upload, onDelete, columns = false }: { schema: SchemaEntry; table: Table; selected: number; write: Write; upload?: Uploader; onDelete: (index: number) => void; columns?: boolean }) {
  const row = table.rows[selected];
  if (!row) return <Empty>This row is not in the table anymore.</Empty>;
  const set = (name: string, value: unknown) => {
    const changes = [{ index: selected, values: { [name]: value } }];
    quiet(write((t) => updateRows(schema.excel!, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  const wide = (f: Field) => !columns || ["text", "file", "url", "image", "string"].includes(f.type);
  return (
    <div className="flex flex-col gap-5">
      <div className={`grid gap-x-4 gap-y-4 ${columns ? "grid-cols-3" : "grid-cols-1"}`}>
        {table.fields.map((f) => (
          <div key={f.name} className={`flex flex-col gap-1.5 ${wide(f) ? "col-span-full" : ""}`}>
            <span className="flex items-center gap-1.5 text-[13px] font-medium text-muted-foreground [&_svg]:size-3.5">
              {TYPE_ICONS[f.type]} {labelOf(f)}
              {f.required && <span className="text-destructive/70">*</span>}
            </span>
            <FieldInput variant="form" field={f} value={row[f.name]} upload={upload} validate={(v) => idError(table, selected, v)} onChange={(v) => set(f.name, v)} />
            {f.description && <span className="text-[12px] text-muted-foreground">{f.description}</span>}
          </div>
        ))}
      </div>
      <button type="button" className="btn ml-auto text-destructive" onClick={() => confirm("Delete this row from the Excel file?") && onDelete(selected)}>
        <Trash2Icon /> Delete the row
      </button>
    </div>
  );
}

/* ----------------------------------- AI ------------------------------------ */

const cell = z.union([z.string(), z.number(), z.boolean(), z.array(z.string()), z.null()]);
const filterArg = z.object({ field: z.string(), op: z.enum(Object.keys(OPS) as [Filter["op"], ...Filter["op"][]]), value: z.string().optional(), values: z.array(z.string()).optional() });

/** The fields for the AI: name, type, rules, choices. */
export const fieldsText = (fields: Field[]) =>
  fields
    .map((f) => `- ${f.name} (${TYPE_LABELS[f.type]}${f.type === "rating" ? ` 1-${f.max ?? 5}` : ""}${f.required ? ", required" : ""}${f.multiple ? ", several values" : ""})${f.label ? ` label "${f.label}"` : ""}${f.options?.length ? ` choices: ${f.options.map((o) => o.value).join(", ")}` : ""}${f.description ? ` — ${f.description}` : ""}`)
    .join("\n");

function excelAgent({ schema, table, shown, checked, view, setView, openRow, write }: {
  schema: SchemaEntry;
  table: Table;
  /** Excel indexes of the rows the view shows, in its order. */
  shown: () => number[];
  checked: number[];
  view: TableView;
  setView: (v: Partial<TableView>) => void;
  openRow: (index: number) => void;
  write: Write;
}): Agent {
  const excel = schema.excel!;
  const idField = table.fields.find(isIdField);
  /** Row given by its id (TASK_001) or its index. */
  const indexOf = (r: { index?: number; id?: string }) => {
    if (r.id === undefined) {
      if (r.index === undefined) throw new Error("Give the row id (or index).");
      return r.index;
    }
    const index = idField ? table.rows.findIndex((x) => String(x[idField.name] ?? "").toLowerCase() === r.id!.toLowerCase()) : -1;
    if (index < 0) throw new Error(`No row with id ${r.id}`);
    return index;
  };
  const pick = (index: number) => ({ index, ...Object.fromEntries(table.fields.map((f) => [f.name, table.rows[index]?.[f.name]])) });
  const flat = (r: Record<string, unknown>) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, Array.isArray(v) ? v.join("; ") : v]));
  const checkFields = (names: string[]) => {
    const unknown = names.filter((n) => !table.fields.some((f) => f.name === n));
    if (unknown.length) throw new Error(`Unknown field(s): ${unknown.join(", ")}. Fields: ${table.fields.map((f) => f.name).join(", ")}.`);
  };
  return {
    instructions: `You manage the CONTENT of an Excel table on SharePoint ("${schema.title}"): you read, find, add, change and delete rows, and change what the user sees (filters, sorts, layout). You never change the columns (that is the Schema editor's job: tell the user). Rows have ${idField ? `an id (column "${idField.name}", e.g. ${idPrefix(excel.table)}_001, given by the app; use it to name rows to the user and in tool calls; change one with change_row_id ONLY when the user explicitly asks to change that id) and ` : ""}an "index" (position in the Excel table). Long text holds markdown; choice fields only accept their listed values (several values: a list); ratings are whole numbers; dates are YYYY-MM-DD; links, images and files are URLs.
Do the whole request in one go; for many rows, change them all in one update_rows call.`,
    context: async () =>
      withContexts(
        [`Excel table "${excel.table}" of "${excel.name}" (schema "${schema.title}")`, `Fields:\n${fieldsText(table.fields)}`, schema.form && `Microsoft Form "${schema.form.title}" fills this table.`],
        [schema.contextIds],
      ),
    // What the user sees now (the "list" of this screen), read again before each model call.
    state: async () =>
      [
        checked.length
          ? `THE USER SELECTED ${checked.length} ROW(S): indexes ${checked.join(", ")}. Act on these rows only (read them with read_rows), unless the user clearly asks otherwise.`
          : "No row selected: the request concerns all the rows.",
        `${table.rows.length} rows; the user's view shows ${shown().length} (${viewText(view, table.fields)}).`,
        `First rows of the view:\n${rowsText(shown().slice(0, 8).map(pick).map(flat), 8)}`,
      ].join("\n\n"),
    tools: [
      tool({
        name: "read_rows",
        description: "Reads rows with their index: the selected rows when the user selected some, else those of the user's view; all=true for every row.",
        args: z.object({ all: z.boolean().optional(), offset: z.number().optional(), limit: z.number().optional() }),
        readOnly: true,
        run: async ({ all, offset = 0, limit = 30 }) => {
          const indexes = all ? table.rows.map((_, i) => i) : checked.length ? checked : shown();
          return rowsText(indexes.slice(offset, offset + limit).map(pick).map(flat), limit);
        },
      }),
      tool({
        name: "find_rows",
        description: 'Finds rows (all of them, whatever the view) matching every filter, e.g. [{"field":"status","op":"is","values":["Open"]}]. Ops: contains, is, isNot, empty, notEmpty, gt, lt, before, after.',
        args: z.object({ filters: z.array(filterArg), limit: z.number().optional() }),
        readOnly: true,
        run: async ({ filters, limit = 30 }) => {
          checkFields(filters.map((f) => f.field));
          const found = table.rows.map((row, index) => ({ row, index })).filter(({ row }) => filters.every((f) => matches(row, f as Filter, table.fields)));
          return `${found.length} row(s).\n${rowsText(found.slice(0, limit).map(({ index }) => flat(pick(index))), limit)}`;
        },
      }),
      tool({
        name: "update_rows",
        description: "Changes values of rows (only the given fields; never the id: see change_row_id). Each row by its id (or index).",
        args: z.object({ rows: z.array(z.object({ id: z.string().optional(), index: z.number().optional(), values: z.record(z.string(), cell) })) }),
        run: async ({ rows: given }) => {
          const changes = given.map((r) => ({ index: indexOf(r), values: Object.fromEntries(Object.entries(r.values).filter(([k]) => k !== idField?.name)) }));
          await write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) });
          return `${changes.length} row(s) updated.`;
        },
      }),
      tool({
        name: "add_rows",
        description: "Adds rows at the end of the table (their ids are given automatically). Required fields must be given.",
        args: z.object({ rows: z.array(z.record(z.string(), cell)) }),
        run: async ({ rows: added }) => {
          await write((t) => addRows(excel, t, added), { refetch: true, optimistic: (t) => withRows(t, added) });
          return `${added.length} row(s) added.`;
        },
      }),
      tool({
        name: "delete_rows",
        description: "Deletes rows by id (or index). Only when the user asks to delete them.",
        args: z.object({ ids: z.array(z.string()).optional(), indexes: z.array(z.number()).optional() }),
        run: async ({ ids = [], indexes: given = [] }) => {
          const indexes = [...new Set([...ids.map((id) => indexOf({ id })), ...given])];
          await write(() => deleteRows(excel, indexes), { optimistic: (t) => withoutRows(t, indexes) });
          return `${indexes.length} row(s) deleted.`;
        },
      }),
      tool({
        name: "set_view",
        description: 'Changes what the user sees (only the given parts): search, filters, sorts, layout ("table", "detail": a list with the open row, "kanban": columns by a choice field = group), grouping, and opens a row. Filter ops as in find_rows; choice fields: op "is" / "isNot" with "values".',
        args: z.object({
          search: z.string().optional(),
          filters: z.array(filterArg).optional(),
          sorts: z.array(z.object({ field: z.string(), desc: z.boolean().optional() })).optional(),
          layout: z.enum(["table", "detail", "kanban"]).optional(),
          group: z.string().nullable().optional().describe("Single choice field giving the kanban's columns (null: the first one)"),
          open_row: z.string().optional().describe("Id of the row to show with all its fields"),
        }),
        run: async ({ search, filters, sorts, layout, group, open_row }) => {
          checkFields([...(filters ?? []), ...(sorts ?? [])].map((x) => x.field).concat(group ? [group] : []));
          setView({ ...(search !== undefined && { search }), ...(filters && { filters: filters as Filter[] }), ...(sorts && { sorts }), ...(layout && { layout: layout === "table" ? undefined : layout }), ...(group !== undefined && { group: group ?? undefined }) });
          if (open_row !== undefined) openRow(indexOf({ id: open_row }));
          return "View updated.";
        },
      }),
      ...(idField
        ? [
            tool({
              name: "change_row_id",
              description: "Changes the id of a row. ONLY when the user explicitly asks to change / rename that id. The new id must be unique (case ignored) and not empty.",
              args: z.object({ id: z.string().describe("Current id"), new_id: z.string() }),
              run: async ({ id, new_id }) => {
                const index = indexOf({ id });
                const error = idError(table, index, new_id);
                if (error) throw new Error(error);
                const changes = [{ index, values: { [idField.name]: new_id.trim() } }];
                await write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) });
                return `Id ${id} changed to ${new_id.trim()}.`;
              },
            }),
          ]
        : []),
    ],
  };
}
