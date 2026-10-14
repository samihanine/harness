import { useEffect, useMemo, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { z } from "zod";
import {
  AlignLeftIcon,
  ArrowDownIcon,
  ArrowLeftIcon,
  ArrowRightIcon,
  ArrowUpIcon,
  CalendarIcon,
  CheckSquareIcon,
  ChevronDownIcon,
  DownloadIcon,
  ExternalLinkIcon,
  EyeOffIcon,
  FilterIcon,
  HashIcon,
  ImageIcon,
  ListIcon,
  MaximizeIcon,
  PlusIcon,
  RefreshCwIcon,
  SearchIcon,
  SlidersHorizontalIcon,
  Trash2Icon,
  TypeIcon,
  XIcon,
} from "lucide-react";
import { Chat } from "@/components/chat";
import { DATASET_TABS, visible as visibleTabs } from "@/config";
import { Picker } from "@/components/picker";
import { FieldInput, type Uploader } from "@/components/field-input";
import { Popover } from "@/components/popover";
import { Empty, SidePanel } from "@/components/side-panel";
import { tool, type Agent } from "@/agent/loop";
import { addRows, deleteRows, idError, resolveFolder, saveFields, updateRows, uploadImage, withRows, withUpdates, withoutRows, type Table } from "@/lib/excel";
import { fieldsFromForm, readFormSchema } from "@/lib/forms";
import { addExcel } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { rowsText } from "@/lib/pbi";
import { COLORS, FIELD_TYPES, TINTS, TYPE_LABELS, idPrefix, isIdField, labelOf, type Field, type FieldType, type Row } from "@/lib/schema";
import { patch, useCollection, type ExcelEntry } from "@/lib/store";
import { EMPTY_VIEW, NO_VALUE, OPS, OPS_OF, apply, filterLabel, groupRows, readView, shownFields, viewText, type Filter, type TableView } from "@/lib/table-view";
import { last, useLast } from "@/lib/last";
import { useTable } from "@/lib/use-table";

export const Route = createFileRoute("/datasets")({
  validateSearch: (s: Record<string, unknown>) => ({ excel: (s.excel as string) || undefined }),
  component: Datasets,
});

type Write = ReturnType<typeof useTable>["write"];
const quiet = (p: Promise<unknown>) => void p.catch(() => undefined); // errors are shown by the page

export const TYPE_ICONS: Record<FieldType, ReactNode> = {
  string: <TypeIcon />,
  text: <AlignLeftIcon />,
  number: <HashIcon />,
  integer: <HashIcon />,
  boolean: <CheckSquareIcon />,
  date: <CalendarIcon />,
  option: <ListIcon />,
  image: <ImageIcon />,
};

/* ---------------------------------- Views ---------------------------------- */

const ALL = "all";

/**
 * The views of a file as tabs: "All rows" (kept in this browser) and the named views (kept with the
 * file in the library). The open view is changed in place; "+" saves it under a new name.
 */
function useViews(excel?: ExcelEntry) {
  const key = `table-view:${excel?.id}`;
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
  const saved = excel?.savedViews ?? [];
  const current = saved.find((v) => v.name === active);
  const view = current ? readView(current.view) : own;
  const setActive = (name: string) => {
    last.set(`${key}:active`, name);
    setActiveState(name);
  };
  const update = (change: Partial<TableView>) => {
    const next = { ...view, ...change };
    if (current && excel) return void patch("excels", excel.id, { savedViews: saved.map((v) => (v.name === current.name ? { ...v, view: next } : v)) });
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
    saveAs: (name: string) => excel && void patch("excels", excel.id, { savedViews: [...saved.filter((v) => v.name !== name), { name, view }] }).then(() => setActive(name)),
    rename: (from: string, to: string) => excel && void patch("excels", excel.id, { savedViews: saved.map((v) => (v.name === from ? { ...v, name: to } : v)) }).then(() => active === from && setActive(to)),
    remove: (name: string) => excel && void patch("excels", excel.id, { savedViews: saved.filter((v) => v.name !== name) }).then(() => setActive(ALL)),
  };
}

/* ----------------------------------- Page ----------------------------------- */

/** What the Row tab shows: a row of the table (its index), a new row, or nothing. */
type RowSelection = number | "new" | undefined;

function Datasets() {
  const { excel: excelId } = Route.useSearch();
  const navigate = useNavigate({ from: "/datasets" });
  useLast("datasets:excel", excelId, (excel) => navigate({ search: { excel }, replace: true }));
  const excels = useCollection("excels");
  const excel = excels.find((e) => e.id === excelId);
  const { table, error, reload, write } = useTable(excel);
  const views = useViews(excel);
  const { view } = views;
  const [tab, setTab] = useState<string>("ai");
  const [row, setRow] = useState<RowSelection>();
  const [column, setColumn] = useState<string>();
  /** Rows checked by the user (Excel indexes): the AI then works on them only. */
  const [checked, setChecked] = useState<number[]>([]);
  useEffect(() => (setChecked([]), setRow(undefined), setColumn(undefined)), [excelId]);

  const fields = table?.fields ?? [];
  const shown = useMemo(() => (table ? apply(table.rows, view, table.fields) : []), [table, view]);
  const openRow = (index: RowSelection) => (setRow(index), setTab("row"));
  const openColumn = (name?: string) => (setColumn(name), setTab("column"));
  const change = (index: number, name: string, value: unknown) => {
    const changes = [{ index, values: { [name]: value } }];
    quiet(write((t) => updateRows(excel!, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  const remove = (indexes: number[]) => {
    quiet(write(() => deleteRows(excel!, indexes), { optimistic: (t) => withoutRows(t, indexes) }));
    setChecked([]);
    setRow(undefined);
  };
  const saveStructure = (next: Field[]) => write((t) => saveFields(excel!, t, next), { refetch: true });
  // Uploads only once a folder is chosen (never a default folder).
  const upload: Uploader | undefined = excel && {
    save: (file) => uploadImage(excel, file),
    folder: excel.imagesFolder?.name,
    chooseFolder: async (link) => patch("excels", excel.id, { imagesFolder: await resolveFolder(link) }),
  };
  const agent = useMemo(
    () => (excel && table ? datasetAgent({ excel, table, shown: () => shown.map((s) => s.index), checked, view, setView: views.update, openRow, write }) : undefined),
    [excel, table, shown, checked, view, write], // eslint-disable-line react-hooks/exhaustive-deps
  );

  return (
    <div className="flex h-full">
      <SidePanel
        active={tab}
        onActive={setTab}
        tabs={visibleTabs([
          { id: "ai", label: "AI", content: agent ? <Chat scope={`dataset:${excelId}`} agent={agent} placeholder="Ask to filter, fix or add rows, change fields…" /> : <Empty>Choose an Excel file.</Empty> },
          { id: "row", label: "Row", content: excel && table ? <RowTab key={String(row)} excel={excel} table={table} selected={row} write={write} upload={upload} onCreated={(index) => setRow(index)} onDelete={(index) => remove([index])} /> : <Empty>Choose an Excel file.</Empty> },
          { id: "column", label: "Column", content: excel && table ? <ColumnTab key={column ?? ""} excel={excel} table={table} name={column} view={view} setView={views.update} onSelect={openColumn} onSave={saveStructure} /> : <Empty>Choose an Excel file.</Empty> },
        ], DATASET_TABS)}
      />
      <div className="flex min-w-0 flex-1 flex-col">
        {/* File and its views */}
        <div className="flex h-12 shrink-0 items-end gap-1 border-b bg-card px-3">
          <Picker
            className="mb-2 mr-2 w-56"
            items={excels.map((e) => ({ id: e.id, label: e.name, hint: e.table }))}
            value={excelId}
            onChange={(excel) => navigate({ search: { excel } })}
            placeholder="Choose an Excel file…"
            add={{
              label: "Add an Excel file",
              placeholder: "Sharing link of an .xlsx (SharePoint / OneDrive)",
              extra: { label: "Microsoft Form filling it (optional)", placeholder: "Link of the form" },
              run: async (link, form) => (await addExcel(link, undefined, [], form)).id,
            }}
          />
          {excel && (
            <>
              {[ALL, ...views.names].map((name) => (
                <button
                  key={name}
                  type="button"
                  onClick={() => views.open(name)}
                  onDoubleClick={() => {
                    if (name === ALL) return;
                    const next = prompt("Rename the view", name)?.trim();
                    if (next) views.rename(name, next);
                  }}
                  className={`group flex h-10 items-center gap-1.5 border-b-2 px-2.5 text-[13px] ${name === views.active ? "border-foreground font-medium" : "border-transparent text-muted-foreground hover:text-foreground"}`}
                >
                  {name === ALL ? "All rows" : name}
                  {name !== ALL && name === views.active && (
                    <XIcon className="size-3 opacity-0 group-hover:opacity-60" onClick={(e) => (e.stopPropagation(), confirm(`Delete the view "${name}"?`) && views.remove(name))} />
                  )}
                </button>
              ))}
              <button type="button" className="icon-btn mb-1.5" title="Save the current filters, sorts and columns as a new view" onClick={() => { const name = prompt("Name of the new view")?.trim(); if (name) views.saveAs(name); }}>
                <PlusIcon />
              </button>
              <span className="ml-auto" />
              <button type="button" className="icon-btn mb-1.5" title="Read the file again" onClick={() => void reload()}>
                <RefreshCwIcon />
              </button>
              <a className="icon-btn mb-1.5" title="Open in Excel" href={excel.link} target="_blank" rel="noreferrer">
                <ExternalLinkIcon />
              </a>
            </>
          )}
        </div>
        {excel && table && <Toolbar fields={fields} view={view} setView={views.update} checked={checked} onClear={() => setChecked([])} onDelete={() => confirm(`Delete ${checked.length} row(s) from the Excel file?`) && remove(checked)} onNew={() => openRow("new")} />}
        {error && <p className="border-b bg-red-500/5 px-4 py-1.5 text-[12px] text-destructive">{error}</p>}
        {!excel && <Empty>Add Excel files (with an Excel table) in the Library, or with the picker above.</Empty>}
        {excel && !table && !error && <Empty>Loading…</Empty>}
        {excel && table && (
          <TableGrid
            table={table}
            items={shown}
            view={view}
            setView={views.update}
            upload={upload}
            onChange={change}
            checked={checked}
            setChecked={setChecked}
            selectedRow={typeof row === "number" ? row : undefined}
            selectedColumn={tab === "column" ? column : undefined}
            onOpenRow={openRow}
            onOpenColumn={openColumn}
            onSaveStructure={saveStructure}
          />
        )}
      </div>
    </div>
  );
}

/* ---------------------------------- Toolbar ---------------------------------- */

/** Search, filters and sorts as chips, grouping, shown fields, selection, new row. */
function Toolbar({ fields, view, setView, checked, onClear, onDelete, onNew }: { fields: Field[]; view: TableView; setView: (v: Partial<TableView>) => void; checked: number[]; onClear: () => void; onDelete: () => void; onNew: () => void }) {
  const shown = shownFields(fields, view);
  return (
    <div className="flex min-h-11 shrink-0 flex-wrap items-center gap-1.5 border-b bg-card px-3 py-1.5">
      <label className="flex h-7 items-center gap-1.5 rounded-md border px-2 focus-within:border-ring">
        <SearchIcon className="size-3.5 text-muted-foreground" />
        <input className="w-28 bg-transparent text-[12px] outline-none" placeholder="Search" value={view.search ?? ""} onChange={(e) => setView({ search: e.target.value })} />
      </label>
      {view.filters.map((f, i) => (
        <Popover key={i} width={300} trigger={(open) => (
          <button type="button" onClick={open} className="flex h-7 max-w-72 items-center gap-1 truncate rounded-full border border-blue-200 bg-blue-50 px-2.5 text-[12px] text-blue-800">
            <FilterIcon className="size-3 shrink-0" /> {filterLabel(f, fields)}
          </button>
        )}>
          {() => <FilterEditor fields={fields} filter={f} onChange={(next) => setView({ filters: view.filters.map((x, k) => (k === i ? next : x)) })} onRemove={() => setView({ filters: view.filters.filter((_, k) => k !== i) })} />}
        </Popover>
      ))}
      {view.sorts.map((s, i) => {
        const field = fields.find((f) => f.name === s.field);
        return (
          <span key={i} className="flex h-7 items-center gap-1 rounded-full border border-violet-200 bg-violet-50 pr-1 pl-2.5 text-[12px] text-violet-800">
            <button type="button" className="flex items-center gap-1" title="Reverse the order" onClick={() => setView({ sorts: view.sorts.map((x, k) => (k === i ? { ...x, desc: !x.desc } : x)) })}>
              {s.desc ? <ArrowDownIcon className="size-3" /> : <ArrowUpIcon className="size-3" />} {field ? labelOf(field) : s.field}
            </button>
            <XIcon className="size-3 cursor-pointer opacity-60 hover:opacity-100" onClick={() => setView({ sorts: view.sorts.filter((_, k) => k !== i) })} />
          </span>
        );
      })}
      <Popover width={200} trigger={(open) => (
        <button type="button" className="btn h-7 text-muted-foreground" onClick={open}>
          <PlusIcon /> Filter or sort
        </button>
      )}>
        {(close) => (
          <div className="flex max-h-96 flex-col overflow-y-auto">
            <span className="label px-1 pb-1">Filter by</span>
            {fields.map((f) => (
              <button key={f.name} type="button" className="flex items-center gap-2 rounded px-2 py-1 text-left text-[12px] hover:bg-muted [&_svg]:size-3.5 [&_svg]:text-muted-foreground" onClick={() => (setView({ filters: [...view.filters, { field: f.name, op: OPS_OF[f.type][0] }] }), close())}>
                {TYPE_ICONS[f.type]} {labelOf(f)}
              </button>
            ))}
            <span className="label mt-1 border-t px-1 pt-2 pb-1">Sort by</span>
            {fields.filter((f) => !view.sorts.some((s) => s.field === f.name)).map((f) => (
              <button key={f.name} type="button" className="flex items-center gap-2 rounded px-2 py-1 text-left text-[12px] hover:bg-muted [&_svg]:size-3.5 [&_svg]:text-muted-foreground" onClick={() => (setView({ sorts: [...view.sorts, { field: f.name }] }), close())}>
                <ArrowUpIcon /> {labelOf(f)}
              </button>
            ))}
          </div>
        )}
      </Popover>
      {checked.length > 0 && (
        <span className="flex h-7 items-center gap-1 rounded-full bg-blue-500/10 pr-0.5 pl-2.5 text-[12px] text-blue-800">
          {checked.length} selected
          <button type="button" className="icon-btn size-6" title="Delete the selected rows" onClick={onDelete}>
            <Trash2Icon />
          </button>
          <button type="button" className="icon-btn size-6" title="Clear the selection" onClick={onClear}>
            <XIcon />
          </button>
        </span>
      )}
      <span className="ml-auto" />
      <select className="input h-7 w-40 text-[12px]" value={view.group ?? ""} onChange={(e) => setView({ group: e.target.value || undefined })} title="Group the rows by an option field">
        <option value="">No grouping</option>
        {fields.filter((f) => f.type === "option").map((f) => (
          <option key={f.name} value={f.name}>Group by {labelOf(f)}</option>
        ))}
      </select>
      <Popover width={220} trigger={(open) => (
        <button type="button" className="btn h-7" onClick={open} title="Columns shown in this view">
          <SlidersHorizontalIcon /> {shown.length}/{fields.length}
        </button>
      )}>
        {() => (
          <div className="flex max-h-96 flex-col gap-1 overflow-y-auto">
            {fields.map((f) => {
              const on = shown.includes(f);
              return (
                <label key={f.name} className="flex items-center gap-2 text-[12px]">
                  <input type="checkbox" className="accent-foreground" checked={on} onChange={() => setView({ hidden: fields.filter((x) => (x === f ? on : !shown.includes(x))).map((x) => x.name) })} />
                  {labelOf(f)}
                </label>
              );
            })}
            {view.hidden && (
              <button type="button" className="btn mt-1 h-7 w-fit" onClick={() => setView({ hidden: undefined })}>
                Default (first 5)
              </button>
            )}
          </div>
        )}
      </Popover>
      <button type="button" className="btn-primary h-7" onClick={onNew}>
        <PlusIcon /> New
      </button>
    </div>
  );
}

/** One filter: field, operator suited to its type, value(s). */
function FilterEditor({ fields, filter, onChange, onRemove }: { fields: Field[]; filter: Filter; onChange: (f: Filter) => void; onRemove: () => void }) {
  const field = fields.find((f) => f.name === filter.field) ?? fields[0];
  return (
    <div className="flex flex-col gap-1.5">
      <div className="flex gap-1.5">
        <select className="input h-7" value={field.name} onChange={(e) => {
          const next = fields.find((f) => f.name === e.target.value)!;
          onChange({ field: next.name, op: OPS_OF[next.type][0] });
        }}>
          {fields.map((f) => (
            <option key={f.name} value={f.name}>{labelOf(f)}</option>
          ))}
        </select>
        <select className="input h-7 w-32" value={filter.op} onChange={(e) => onChange({ ...filter, op: e.target.value as Filter["op"] })}>
          {OPS_OF[field.type].map((o) => (
            <option key={o} value={o}>{OPS[o]}</option>
          ))}
        </select>
        <button type="button" className="icon-btn" title="Remove the filter" onClick={onRemove}>
          <XIcon />
        </button>
      </div>
      {!NO_VALUE.includes(filter.op) &&
        (field.type === "option" ? (
          <div className="flex flex-wrap gap-1">
            {field.options?.map((o) => {
              const on = filter.values?.includes(o.value);
              return (
                <button key={o.value} type="button" onClick={() => onChange({ ...filter, values: on ? filter.values!.filter((v) => v !== o.value) : [...(filter.values ?? []), o.value] })} className={`inline-flex h-5 items-center rounded-full px-2 text-[11px] ${TINTS[o.color]} ${on ? "ring-2 ring-foreground/50 ring-offset-1" : "opacity-60 hover:opacity-100"}`}>
                  {o.value}
                </button>
              );
            })}
          </div>
        ) : (
          <input autoFocus className="input h-7" type={field.type === "number" || field.type === "integer" ? "number" : field.type === "date" ? "date" : "text"} placeholder="Value" value={filter.value ?? ""} onChange={(e) => onChange({ ...filter, value: e.target.value })} />
        ))}
    </div>
  );
}

/* ----------------------------------- Table ----------------------------------- */

function TableGrid({ table, items, view, setView, upload, onChange, checked, setChecked, selectedRow, selectedColumn, onOpenRow, onOpenColumn, onSaveStructure }: {
  table: Table;
  items: { row: Row; index: number }[];
  view: TableView;
  setView: (v: Partial<TableView>) => void;
  upload?: Uploader;
  onChange: (index: number, name: string, value: unknown) => void;
  checked: number[];
  setChecked: (c: number[]) => void;
  selectedRow?: number;
  selectedColumn?: string;
  onOpenRow: (index: number | "new") => void;
  onOpenColumn: (name?: string) => void;
  onSaveStructure: (fields: Field[]) => Promise<unknown>;
}) {
  const columns = shownFields(table.fields, view);
  const groupField = table.fields.find((f) => f.name === view.group && f.type === "option");
  const all = items.length > 0 && items.every((i) => checked.includes(i.index));
  // Drag a header onto another to move the column (order saved in the "_schema" sheet).
  const [dragged, setDragged] = useState<string>();
  // Drag a header's right edge to resize the column (saved with the view when released).
  const [resizing, setResizing] = useState<{ name: string; width: number }>();
  const widthOf = (f: Field) => (resizing?.name === f.name ? resizing.width : view.widths?.[f.name]) ?? (isIdField(f) ? 100 : f.type === "text" ? 300 : f.type === "string" ? 200 : f.type === "image" ? 90 : 140);
  const style = (f: Field) => ({ width: widthOf(f), minWidth: widthOf(f), maxWidth: widthOf(f) });
  const resize = (f: Field, e: React.PointerEvent<HTMLElement>) => {
    e.preventDefault();
    e.stopPropagation();
    const start = e.clientX;
    const from = widthOf(f);
    let current = from;
    const move = (m: PointerEvent) => setResizing({ name: f.name, width: (current = Math.max(56, Math.round(from + m.clientX - start))) });
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setResizing(undefined);
      setView({ widths: { ...view.widths, [f.name]: current } });
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };
  const move = (from: string, to: string) => {
    if (from === to) return;
    const next = [...table.fields];
    const [moved] = next.splice(next.findIndex((f) => f.name === from), 1);
    next.splice(next.findIndex((f) => f.name === to) + (table.fields.findIndex((f) => f.name === from) < table.fields.findIndex((f) => f.name === to) ? 1 : 0), 0, moved);
    quiet(onSaveStructure(next));
  };
  return (
    <div className="min-h-0 flex-1 overflow-auto bg-card">
      <table className="w-max min-w-full border-separate border-spacing-0 text-[13px]">
        <thead className="sticky top-0 z-10 bg-card">
          <tr>
            <th className="w-9 border-r border-b">
              <span className="flex justify-center">
                <input type="checkbox" className="size-3.5 accent-foreground" title="Select all the rows shown" checked={all} onChange={() => setChecked(all ? [] : items.map((i) => i.index))} />
              </span>
            </th>
            {columns.map((f) => {
              const sort = view.sorts.find((s) => s.field === f.name);
              return (
                <th
                  key={f.name}
                  style={style(f)}
                  draggable={!resizing}
                  onDragStart={(e) => (e.dataTransfer.setData("text/plain", f.name), setDragged(f.name))}
                  onDragOver={(e) => dragged && e.preventDefault()}
                  onDrop={(e) => (e.preventDefault(), dragged && move(dragged, f.name), setDragged(undefined))}
                  onDragEnd={() => setDragged(undefined)}
                  className={`relative border-r border-b p-0 text-left font-normal ${selectedColumn === f.name ? "bg-accent" : ""} ${dragged === f.name ? "opacity-40" : ""}`}
                >
                  <button type="button" title={f.description ?? "Column settings (drag to move)"} onClick={() => onOpenColumn(f.name)} className="flex h-9 w-full items-center gap-1.5 px-3 text-[12px] text-muted-foreground hover:bg-muted/50 hover:text-foreground [&_svg]:size-3.5 [&_svg]:shrink-0">
                    {TYPE_ICONS[f.type]}
                    <span className="truncate">{labelOf(f)}</span>
                    {f.required && <span className="text-destructive/70">*</span>}
                    {sort && (sort.desc ? <ArrowDownIcon /> : <ArrowUpIcon />)}
                  </button>
                  <div role="separator" aria-orientation="vertical" title="Drag to resize" onPointerDown={(e) => resize(f, e)} onDragStart={(e) => e.preventDefault()} className={`absolute inset-y-0 -right-1 z-10 w-2 cursor-col-resize hover:bg-ring/40 ${resizing?.name === f.name ? "bg-ring/50" : ""}`} />
                </th>
              );
            })}
            <th className="w-10 border-b p-0">
              <AddField table={table} onSave={onSaveStructure} />
            </th>
          </tr>
        </thead>
        <tbody>
          {groupRows(items, groupField).map((g) => (
            <Group key={g.key} label={groupField ? g.label : undefined} color={g.color} count={g.items.length} span={columns.length + 2}>
              {g.items.map(({ row, index }) => (
                <tr key={index} className={`group ${selectedRow === index ? "bg-accent" : checked.includes(index) ? "bg-blue-500/5" : "hover:bg-muted/30"}`}>
                  <td className="border-r border-b align-middle">
                    <span className="flex justify-center">
                      <input type="checkbox" className="size-3.5 accent-foreground" title="Select this row" checked={checked.includes(index)} onChange={() => setChecked(checked.includes(index) ? checked.filter((i) => i !== index) : [...checked, index])} />
                    </span>
                  </td>
                  {columns.map((f, c) => (
                    <td key={f.name} style={style(f)} className="relative border-r border-b p-0 align-top">
                      <FieldInput variant="cell" field={f} value={row[f.name]} upload={upload} validate={(v) => idError(table, index, v)} onChange={(v) => onChange(index, f.name, v)} />
                      {c === 0 && (
                        <button type="button" title="All the fields of this row" onClick={() => onOpenRow(index)} className="absolute top-1.5 right-1 hidden h-6 items-center gap-1 rounded border bg-card px-1.5 text-[11px] text-muted-foreground shadow-sm group-hover:flex hover:text-foreground">
                          <MaximizeIcon className="size-3" /> Open
                        </button>
                      )}
                    </td>
                  ))}
                  <td className="border-b" />
                </tr>
              ))}
            </Group>
          ))}
        </tbody>
      </table>
      {items.length === 0 && <p className="p-6 text-center text-[12px] text-muted-foreground">No row matches this view.</p>}
      <button type="button" className="flex w-full items-center gap-1.5 px-3 py-2 text-left text-[12px] text-muted-foreground hover:bg-muted/50 [&_svg]:size-3.5" onClick={() => onOpenRow("new")}>
        <PlusIcon /> New row
      </button>
    </div>
  );
}

/** A group of rows (grouping by an option field), foldable; no header without grouping. */
function Group({ label, color, count, span, children }: { label?: string; color?: keyof typeof TINTS; count: number; span: number; children: ReactNode }) {
  const [folded, setFolded] = useState(false);
  return (
    <>
      {label !== undefined && (
        <tr>
          <td colSpan={span} className="border-b bg-muted/40 px-2 py-1.5">
            <button type="button" className="flex items-center gap-2 text-[12px] font-medium" onClick={() => setFolded(!folded)}>
              <ChevronDownIcon className={`size-3.5 transition-transform ${folded ? "-rotate-90" : ""}`} />
              <span className={`rounded-full px-2 ${color ? TINTS[color] : "bg-muted"}`}>{label}</span>
              <span className="text-muted-foreground">{count}</span>
            </button>
          </td>
        </tr>
      )}
      {!folded && children}
    </>
  );
}

/** The "+" column: a new field (a new Excel column) with its type. */
function AddField({ table, onSave }: { table: Table; onSave: (fields: Field[]) => Promise<unknown> }) {
  const [name, setName] = useState("");
  const [type, setType] = useState<FieldType>("string");
  const [error, setError] = useState<string>();
  return (
    <Popover width={260} trigger={(open) => (
      <button type="button" title="Add a field (new Excel column)" onClick={open} className="flex h-9 w-10 items-center justify-center text-muted-foreground hover:bg-muted/50 hover:text-foreground">
        <PlusIcon className="size-4" />
      </button>
    )}>
      {(close) => (
        <div className="flex flex-col gap-2">
          <input autoFocus className="input h-7" placeholder="Column name" value={name} onChange={(e) => setName(e.target.value)} />
          <div className="grid grid-cols-2 gap-1">
            {FIELD_TYPES.map((t) => (
              <button key={t} type="button" onClick={() => setType(t)} className={`flex items-center gap-1.5 rounded-md border px-2 py-1 text-[12px] [&_svg]:size-3.5 ${t === type ? "border-foreground bg-muted" : "hover:bg-muted/50"}`}>
                {TYPE_ICONS[t]} {TYPE_LABELS[t]}
              </button>
            ))}
          </div>
          {error && <p className="text-[11px] text-destructive">{error}</p>}
          <button
            type="button"
            className="btn-primary h-7"
            disabled={!name.trim()}
            onClick={() => onSave([...table.fields, { name: name.trim(), type, options: type === "option" ? [] : undefined }]).then(() => (setName(""), setError(undefined), close()), (e) => setError(errorText(e)))}
          >
            Add the field
          </button>
        </div>
      )}
    </Popover>
  );
}

/* ---------------------------------- Row tab ---------------------------------- */

/** Every field of the open row (or of a new one), as a form. */
function RowTab({ excel, table, selected, write, upload, onCreated, onDelete }: { excel: ExcelEntry; table: Table; selected: RowSelection; write: Write; upload?: Uploader; onCreated: (index: number) => void; onDelete: (index: number) => void }) {
  const [draft, setDraft] = useState<Row>({});
  const [error, setError] = useState<string>();
  if (selected === undefined) return <Empty>Click "Open" on a row (or New) to see all its fields here.</Empty>;
  const isNew = selected === "new";
  const row = isNew ? draft : table.rows[selected];
  if (!row) return <Empty>This row is not in the table anymore.</Empty>;
  const set = (name: string, value: unknown) => {
    if (isNew) return setDraft((d) => ({ ...d, [name]: value }));
    const changes = [{ index: selected, values: { [name]: value } }];
    quiet(write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  const create = () =>
    write((t) => addRows(excel, t, [draft]), { refetch: true, optimistic: (t) => withRows(t, [draft]) }).then(
      () => onCreated(table.rows.length),
      (e) => setError(errorText(e)),
    );
  return (
    <div className="flex flex-col gap-3 p-3">
      <div className="flex items-center gap-2">
        <span className="text-[13px] font-medium">{isNew ? "New row" : `Row ${selected + 1}`}</span>
        {!isNew && (
          <button type="button" className="btn ml-auto h-7 text-destructive" onClick={() => confirm("Delete this row from the Excel file?") && onDelete(selected)}>
            <Trash2Icon /> Delete
          </button>
        )}
      </div>
      {table.fields.filter((f) => !(isNew && isIdField(f))).map((f) => (
        <div key={f.name} className="flex flex-col gap-1">
          <span className="flex items-center gap-1.5 text-[12px] font-medium [&_svg]:size-3.5 [&_svg]:text-muted-foreground">
            {TYPE_ICONS[f.type]} {labelOf(f)}
            {f.required && <span className="text-destructive/70">*</span>}
          </span>
          <FieldInput variant="form" field={f} value={row[f.name]} upload={upload} validate={isNew ? undefined : (v) => idError(table, selected, v)} onChange={(v) => set(f.name, v)} />
          {f.description && <span className="text-[11px] text-muted-foreground">{f.description}</span>}
        </div>
      ))}
      {error && <p className="text-[12px] break-words text-destructive">{error}</p>}
      {isNew && (
        <button type="button" className="btn-primary w-fit" onClick={() => void create()}>
          <PlusIcon /> Create the row
        </button>
      )}
    </div>
  );
}

/* -------------------------------- Column tab -------------------------------- */

/**
 * The open column's settings (label, type, options, rules, place), saved to the "_schema" sheet; with
 * no column open, the table's structure: its fields, the Microsoft Form filling it, its schema.
 */
function ColumnTab({ excel, table, name, view, setView, onSelect, onSave }: { excel: ExcelEntry; table: Table; name?: string; view: TableView; setView: (v: Partial<TableView>) => void; onSelect: (name?: string) => void; onSave: (fields: Field[]) => Promise<unknown> }) {
  const index = table.fields.findIndex((f) => f.name === name);
  const field = table.fields[index];
  const [draft, setDraft] = useState<Field | undefined>(field);
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  if (!field || !draft) return <TableStructure excel={excel} table={table} onSelect={onSelect} onSave={onSave} />;
  const changed = JSON.stringify(draft) !== JSON.stringify(field);
  const save = (fields: Field[], after?: () => void) => {
    setState({ busy: true });
    onSave(fields).then(() => (setState({}), after?.()), (e) => setState({ error: errorText(e) }));
  };
  const set = (p: Partial<Field>) => setDraft({ ...draft, ...p });
  const hidden = !shownFields(table.fields, view).includes(field);
  return (
    <div className="flex flex-col gap-3 p-3">
      <button type="button" className="flex w-fit items-center gap-1 text-[12px] text-muted-foreground hover:text-foreground" onClick={() => onSelect(undefined)}>
        <ArrowLeftIcon className="size-3" /> All the fields
      </button>
      <div className="flex flex-col gap-1">
        <span className="label">Column (in Excel)</span>
        <input className="input" value={draft.name} disabled title="The Excel column keeps its name; change the label" readOnly />
      </div>
      <label className="flex flex-col gap-1">
        <span className="label">Label</span>
        <input className="input" value={draft.label ?? ""} placeholder={labelOf({ ...draft, label: undefined })} onChange={(e) => set({ label: e.target.value || undefined })} />
      </label>
      <div className="flex flex-col gap-1">
        <span className="label">Type</span>
        <div className="grid grid-cols-2 gap-1">
          {FIELD_TYPES.map((t) => (
            <button key={t} type="button" onClick={() => set({ type: t, options: t === "option" ? draft.options ?? [] : undefined })} className={`flex items-center gap-1.5 rounded-md border px-2 py-1 text-[12px] [&_svg]:size-3.5 ${t === draft.type ? "border-foreground bg-muted" : "hover:bg-muted/50"}`}>
              {TYPE_ICONS[t]} {TYPE_LABELS[t]}
            </button>
          ))}
        </div>
      </div>
      {draft.type === "option" && (
        <div className="flex flex-col gap-1">
          <span className="label">Options (their order sorts the column)</span>
          {(draft.options ?? []).map((o, k) => (
            <div key={k} className="flex items-center gap-1.5">
              <select className={`h-6 rounded-full px-1 text-[11px] outline-none ${TINTS[o.color]}`} value={o.color} onChange={(e) => set({ options: draft.options!.map((x, j) => (j === k ? { ...x, color: e.target.value as never } : x)) })}>
                {COLORS.map((c) => (
                  <option key={c}>{c}</option>
                ))}
              </select>
              <input className="input h-7 flex-1" value={o.value} onChange={(e) => set({ options: draft.options!.map((x, j) => (j === k ? { ...x, value: e.target.value } : x)) })} />
              <button type="button" className="icon-btn" title="Remove the option" onClick={() => set({ options: draft.options!.filter((_, j) => j !== k) })}>
                <XIcon />
              </button>
            </div>
          ))}
          <button type="button" className="btn h-7 w-fit" onClick={() => set({ options: [...(draft.options ?? []), { value: "New", color: COLORS[((draft.options?.length ?? 0) + 1) % COLORS.length] }] })}>
            <PlusIcon /> Option
          </button>
          <label className="flex items-center gap-2 text-[12px]">
            <input type="checkbox" className="accent-foreground" checked={!!draft.multiple} onChange={(e) => set({ multiple: e.target.checked || undefined })} /> Several values per row
          </label>
        </div>
      )}
      <label className="flex flex-col gap-1">
        <span className="label">Description</span>
        <input className="input" placeholder="Shown under the field (and to the AI)" value={draft.description ?? ""} onChange={(e) => set({ description: e.target.value || undefined })} />
      </label>
      <label className="flex items-center gap-2 text-[12px]">
        <input type="checkbox" className="accent-foreground" disabled={isIdField(draft)} checked={!!draft.required} onChange={(e) => set({ required: e.target.checked || undefined })} /> Required
      </label>
      {changed && (
        <div className="flex gap-1.5">
          <button type="button" className="btn-primary" disabled={state.busy} onClick={() => save(table.fields.map((f, k) => (k === index ? { ...draft, options: draft.type === "option" ? (draft.options ?? []).filter((o) => o.value.trim()) : undefined } : f)))}>
            {state.busy ? "Saving…" : "Save to the Excel file"}
          </button>
          <button type="button" className="btn" onClick={() => setDraft(field)}>
            Cancel
          </button>
        </div>
      )}
      {state.error && <p className="text-[12px] break-words text-destructive">{state.error}</p>}

      <div className="flex flex-col gap-1.5 border-t pt-3">
        <span className="label">In this view</span>
        <div className="grid grid-cols-2 gap-1 text-[12px]">
          <button type="button" className="btn h-7" onClick={() => setView({ sorts: [{ field: field.name }] })}>
            <ArrowUpIcon /> Sort A → Z
          </button>
          <button type="button" className="btn h-7" onClick={() => setView({ sorts: [{ field: field.name, desc: true }] })}>
            <ArrowDownIcon /> Sort Z → A
          </button>
          <button type="button" className="btn h-7" onClick={() => setView({ filters: [...view.filters, { field: field.name, op: OPS_OF[field.type][0] }] })}>
            <FilterIcon /> Filter
          </button>
          <button type="button" className="btn h-7" onClick={() => setView({ hidden: table.fields.filter((f) => (f === field ? !hidden : !shownFields(table.fields, view).includes(f))).map((f) => f.name) })}>
            <EyeOffIcon /> {hidden ? "Show" : "Hide"}
          </button>
        </div>
      </div>
      <div className="flex gap-1.5 border-t pt-3">
        <button type="button" className="btn h-7" title="Move left" disabled={index === 0 || state.busy} onClick={() => save(swap(table.fields, index, index - 1))}>
          <ArrowLeftIcon />
        </button>
        <button type="button" className="btn h-7" title="Move right" disabled={index === table.fields.length - 1 || state.busy} onClick={() => save(swap(table.fields, index, index + 1))}>
          <ArrowRightIcon />
        </button>
        {!isIdField(field) && (
          <button type="button" className="btn ml-auto h-7 text-destructive" title="The Excel column is kept, only hidden from the app" onClick={() => confirm(`Remove "${labelOf(field)}" from the fields? (the Excel column is kept)`) && save(table.fields.filter((_, k) => k !== index), () => onSelect(undefined))}>
            <Trash2Icon /> Remove the field
          </button>
        )}
      </div>
    </div>
  );
}

const swap = <T,>(list: T[], a: number, b: number) => list.map((x, k) => (k === a ? list[b] : k === b ? list[a] : x));

/** No column open: the fields (each opens its settings), adding one, the form filling the table, the schema. */
function TableStructure({ excel, table, onSelect, onSave }: { excel: ExcelEntry; table: Table; onSelect: (name: string) => void; onSave: (fields: Field[]) => Promise<unknown> }) {
  const json = JSON.stringify(schemaOf(excel, table), null, 2);
  const download = () => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([json], { type: "application/json" }));
    a.download = `${excel.table}.schema.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  };
  return (
    <div className="flex flex-col gap-3 p-3">
      <p className="text-[12px] text-muted-foreground">Click a column header (or a field below) to change it.</p>
      {!table.hasSchema && <p className="rounded-md bg-amber-500/10 p-2 text-[12px] text-amber-800">The types are guessed from the data: saving a field writes them to a hidden "_schema" sheet.</p>}
      <ul className="flex flex-col divide-y rounded-lg border">
        {table.fields.map((f) => (
          <li key={f.name}>
            <button type="button" onClick={() => onSelect(f.name)} className="flex w-full items-center gap-2 px-2.5 py-1.5 text-left text-[13px] hover:bg-muted/50 [&_svg]:size-3.5 [&_svg]:text-muted-foreground">
              {TYPE_ICONS[f.type]}
              <span className="min-w-0 flex-1 truncate">{labelOf(f)}</span>
              <span className="text-[11px] text-muted-foreground">{TYPE_LABELS[f.type]}</span>
            </button>
          </li>
        ))}
      </ul>
      <div className="w-fit">
        <AddFieldButton table={table} onSave={onSave} />
      </div>
      <FormLink excel={excel} table={table} onAddFields={(added) => onSave([...table.fields, ...added])} />
      <div className="flex items-center gap-2">
        <span className="label">Schema (JSON)</span>
        <button type="button" className="btn ml-auto h-7" onClick={download}>
          <DownloadIcon /> Download
        </button>
      </div>
    </div>
  );
}

const AddFieldButton = ({ table, onSave }: { table: Table; onSave: (fields: Field[]) => Promise<unknown> }) => (
  <span className="flex items-center gap-1.5 rounded-md border pr-2 text-[12px]">
    <AddField table={table} onSave={onSave} /> Add a field
  </span>
);

/** The table's schema as JSON (fields, and the questions of its Microsoft Form when one is linked). */
export const schemaOf = (excel: ExcelEntry, table: Table) => ({
  name: excel.name,
  table: excel.table,
  link: excel.link,
  fields: table.fields,
  form: excel.form ?? null,
});

/** The Microsoft Form filling the table (optional): link, read again, remove, make the table's fields from it. */
function FormLink({ excel, table, onAddFields }: { excel: ExcelEntry; table: Table; onAddFields: (fields: Field[]) => Promise<unknown> }) {
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string; done?: string }>({});
  const missing = excel.form ? fieldsFromForm(excel.form, table.fields) : [];
  const addFields = () => {
    if (!confirm(`Add ${missing.length} field(s) (new Excel columns)?\n\n${missing.map((f) => `• ${f.name} — ${TYPE_LABELS[f.type]}${f.multiple ? ", multiple" : ""}${f.options?.length ? ` (${f.options.length} options)` : ""}`).join("\n")}`)) return;
    setState({ busy: true });
    onAddFields(missing).then(() => setState({ done: `${missing.length} field(s) added.` }), (e) => setState({ error: errorText(e) }));
  };
  const read = (url: string) => {
    setState({ busy: true });
    readFormSchema(url).then(
      (form) => patch("excels", excel.id, { form }).then(() => (setState({}), setLink(""))),
      (e) => setState({ error: errorText(e) }),
    );
  };
  return (
    <div className="flex flex-col gap-1.5 rounded-lg border p-2.5">
      <span className="label">Microsoft Form</span>
      {excel.form ? (
        <div className="flex items-center gap-1">
          <a className="min-w-0 flex-1 truncate font-medium underline" href={excel.form.url} target="_blank" rel="noreferrer">
            {excel.form.title || "Untitled form"}
          </a>
          <span className="shrink-0 text-[11px] text-muted-foreground">{excel.form.questions.length} questions</span>
          <button type="button" className="icon-btn" title="Read the form again" disabled={state.busy} onClick={() => read(excel.form!.url)}>
            <RefreshCwIcon />
          </button>
          <button type="button" className="icon-btn" title="Unlink the form" onClick={() => void patch("excels", excel.id, { form: undefined })}>
            <XIcon />
          </button>
        </div>
      ) : (
        <div className="flex gap-1.5">
          <input className="input h-7" placeholder="Link of the form (optional)" value={link} onChange={(e) => setLink(e.target.value)} onKeyDown={(e) => e.key === "Enter" && link.trim() && read(link)} />
          <button type="button" className="btn-primary h-7" disabled={state.busy || !link.trim()} onClick={() => read(link)}>
            {state.busy ? "Reading…" : "Link"}
          </button>
        </div>
      )}
      {excel.form && (
        <button type="button" className="btn h-7 w-fit" disabled={state.busy || !missing.length} title="One field per question that has none yet: type, options, multiple, required" onClick={addFields}>
          <PlusIcon /> {missing.length ? `Create the fields from the form (${missing.length})` : "Every question has its field"}
        </button>
      )}
      {state.done && <p className="text-[11px] text-green-700">{state.done}</p>}
      {state.error && <p className="text-[11px] break-words text-destructive">{state.error}</p>}
    </div>
  );
}

/* ----------------------------------- AI ------------------------------------ */

const cell = z.union([z.string(), z.number(), z.boolean(), z.array(z.string()), z.null()]);

function datasetAgent({ excel, table, shown, checked, view, setView, openRow, write }: {
  excel: ExcelEntry;
  table: Table;
  /** Excel indexes of the rows the view shows, in its order. */
  shown: () => number[];
  checked: number[];
  view: TableView;
  setView: (v: Partial<TableView>) => void;
  openRow: (index: number) => void;
  write: Write;
}): Agent {
  const idField = table.fields.find(isIdField);
  /** Row given by its id (TASK_001) or its index. */
  const indexOf = (r: { index?: number; id?: string }) => {
    if (r.id === undefined) {
      if (r.index === undefined) throw new Error("Give the row id (or index).");
      return r.index;
    }
    const index = idField ? table.rows.findIndex((x) => String(x[idField.name] ?? "") === r.id) : -1;
    if (index < 0) throw new Error(`No row with id ${r.id}`);
    return index;
  };
  const fieldsText = table.fields
    .map((f) => `- ${f.name} (${TYPE_LABELS[f.type]}${f.required ? ", required" : ""}${f.multiple ? ", multiple" : ""})${f.label ? ` label "${f.label}"` : ""}${f.options?.length ? ` options: ${f.options.map((o) => `${o.value} [${o.color}]`).join(", ")}` : ""}${f.description ? ` — ${f.description}` : ""}`)
    .join("\n");
  return {
    instructions: `You manage the content of an Excel table on SharePoint ("${excel.name}"), shown as a table with filters, sorts and grouping. Rows have ${idField ? `an id (column "${idField.name}", e.g. ${idPrefix(excel.table)}_001, given by the app; use it to name rows to the user and in tool calls; change one with change_row_id ONLY when the user explicitly asks to change that id) and ` : ""}an "index" (position in the Excel table). Text fields hold markdown; option fields only accept their listed values (multiple: a list); image fields hold image links; dates are YYYY-MM-DD.`,
    context: async () =>
      [
        `Excel "${excel.name}" (table ${excel.table})${excel.context ? `\nNotes: ${excel.context}` : ""}`,
        `Fields:\n${fieldsText}`,
        excel.form ? `Microsoft Form "${excel.form.title}" fills this table; its questions: ${excel.form.questions.map((q) => `${q.title} (${q.type}${q.multiple ? ", multiple" : ""}${q.required ? ", required" : ""})`).join("; ")}` : "",
      ].filter(Boolean).join("\n\n"),
    state: async () =>
      [
        checked.length
          ? `THE USER SELECTED ${checked.length} ROW(S): indexes ${checked.join(", ")}. Act on these rows only (read them with read_rows), unless the user clearly asks otherwise.`
          : "No row selected: the request concerns all the rows.",
        `${table.rows.length} rows; the user's view shows ${shown().length} (${viewText(view, table.fields)}).`,
      ].join("\n\n"),
    tools: [
      tool({
        name: "read_rows",
        description: "Reads rows with their index: the selected rows when the user selected some, else those of the user's view; all=true for every row.",
        args: z.object({ all: z.boolean().optional(), offset: z.number().optional(), limit: z.number().optional() }),
        readOnly: true,
        run: async ({ all, offset = 0, limit = 30 }) => {
          // Only the fields (hidden Excel columns are not shown).
          const pick = (index: number) => ({ index, ...Object.fromEntries(table.fields.map((f) => [f.name, table.rows[index]?.[f.name]])) });
          const indexes = all ? table.rows.map((_, i) => i) : checked.length ? checked : shown();
          return rowsText(indexes.slice(offset, offset + limit).map(pick).map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, Array.isArray(v) ? v.join("; ") : v]))), limit);
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
        description: "Adds rows at the end of the table (their ids are given automatically).",
        args: z.object({ rows: z.array(z.record(z.string(), cell)) }),
        run: async ({ rows: added }) => {
          await write((t) => addRows(excel, t, added), { refetch: true, optimistic: (t) => withRows(t, added) });
          return `${added.length} row(s) added.`;
        },
      }),
      tool({
        name: "delete_rows",
        description: "Deletes rows by id (or index).",
        args: z.object({ ids: z.array(z.string()).optional(), indexes: z.array(z.number()).optional() }),
        run: async ({ ids = [], indexes: given = [] }) => {
          const indexes = [...new Set([...ids.map((id) => indexOf({ id })), ...given])];
          await write(() => deleteRows(excel, indexes), { optimistic: (t) => withoutRows(t, indexes) });
          return `${indexes.length} row(s) deleted.`;
        },
      }),
      tool({
        name: "set_view",
        description: 'Changes what the user sees (only the given parts): search, filters, sorts, grouping, and opens a row. Filter ops: contains, is, isNot, empty, notEmpty, gt, lt, before, after (YYYY-MM-DD), true, false; option fields: op "is" / "isNot" with "values".',
        args: z.object({
          search: z.string().optional(),
          filters: z.array(z.object({ field: z.string(), op: z.enum(Object.keys(OPS) as [Filter["op"], ...Filter["op"][]]), value: z.string().optional(), values: z.array(z.string()).optional() })).optional(),
          sorts: z.array(z.object({ field: z.string(), desc: z.boolean().optional() })).optional(),
          group: z.string().nullable().optional().describe("Option field to group the rows by (null: no grouping)"),
          open_row: z.number().optional().describe("Index of the row to show with all its fields"),
        }),
        run: async ({ search, filters, sorts, group, open_row }) => {
          const unknown = [...(filters ?? []), ...(sorts ?? [])].map((x) => x.field).filter((name) => !table.fields.some((f) => f.name === name));
          if (unknown.length) throw new Error(`Unknown field(s): ${unknown.join(", ")}. Fields: ${table.fields.map((f) => f.name).join(", ")}.`);
          setView({ ...(search !== undefined && { search }), ...(filters && { filters }), ...(sorts && { sorts }), ...(group !== undefined && { group: group ?? undefined }) });
          if (open_row !== undefined) openRow(open_row);
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
