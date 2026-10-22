import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import {
  AlignLeftIcon,
  ArrowDownIcon,
  ArrowUpDownIcon,
  ArrowUpIcon,
  BookTextIcon,
  BookmarkIcon,
  BookmarkPlusIcon,
  CalendarIcon,
  ChevronDownIcon,
  ClipboardListIcon,
  ExternalLinkIcon,
  FileIcon,
  FileSpreadsheetIcon,
  FolderIcon,
  HashIcon,
  ImageIcon,
  LinkIcon,
  ListIcon,
  Maximize2Icon,
  PlusIcon,
  RefreshCwIcon,
  SaveIcon,
  SearchIcon,
  Settings2Icon,
  SlidersHorizontalIcon,
  StarIcon,
  Table2Icon,
  Trash2Icon,
  TypeIcon,
  UploadIcon,
  XIcon,
} from "lucide-react";
import { ContextLinks } from "@/components/context-links";
import { AddDialog, EntityMenu, withToast } from "@/components/entity";
import { Badge, FieldInput, type Uploader } from "@/components/field-input";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { DropdownMenuItem } from "@/components/ui/dropdown-menu";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from "@/components/ui/sheet";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { EntityPicker, Empty, Section, StatusText, SubTabs, Workspace } from "@/components/workspace";
import { EXCEL_UPLOAD } from "@/config";
import { excelData } from "@/agent/excel-tools";
import type { Agent } from "@/agent/loop";
import { schemaAgent } from "@/agent/schema-tools";
import { addRows, applySchemaToExcel, deleteRows, idError, resolveFolder, updateRows, uploadFileTo, withRows, withUpdates, withoutRows, type Table } from "@/lib/excel";
import { applyForm, designUrl, planForm, type FormPlan } from "@/lib/forms";
import { addSchema, downloadText, linkForm, saveFields, setBuiltInExcel } from "@/lib/library";
import { useLast } from "@/lib/last";
import { errorText } from "@/lib/ms";
import { COLORS, FIELD_TYPES, SWATCHES, TYPE_LABELS, fieldsErrors, isIdField, labelOf, type Field, type FieldType, type FormSection, type Option, type Row } from "@/lib/schema";
import { patch, remove, useCollection, type SchemaEntry } from "@/lib/store";
import { EMPTY_VIEW, NO_VALUE, OPS, OPS_OF, apply, filterLabel, readView, type Filter, type TableView } from "@/lib/table-view";
import { useTable, type Write } from "@/lib/use-table";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/excel")({
  validateSearch: (s: Record<string, unknown>): { schema?: string; tab?: string } => ({ schema: (s.schema as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Excel,
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

/* ------------------------------ View and bookmarks ------------------------------ */

/**
 * What the Data tab shows (search, filters, sorts), kept in this browser, and the bookmarks: named
 * views kept with the schema (downloaded and imported with it). A bookmark is applied, not edited in place.
 */
function useView(schema?: SchemaEntry) {
  const key = `excel-view:${schema?.id}`;
  const [view, setViewState] = useState<TableView>(EMPTY_VIEW);
  useEffect(() => {
    try {
      setViewState(readView(JSON.parse(localStorage.getItem(key) ?? "{}")));
    } catch {
      setViewState(EMPTY_VIEW);
    }
  }, [key]);
  const set = (next: TableView) => {
    setViewState(next);
    try {
      localStorage.setItem(key, JSON.stringify(next));
    } catch {}
  };
  const bookmarks = schema?.views ?? [];
  const same = (v: Record<string, unknown>) => {
    const b = readView(v);
    return JSON.stringify([b.search ?? "", b.filters, b.sorts]) === JSON.stringify([view.search ?? "", view.filters, view.sorts]);
  };
  return {
    view,
    update: (change: Partial<TableView>) => set({ ...view, ...change }),
    bookmarks: bookmarks.map((b) => ({ name: b.name, active: same(b.view) })),
    open: (name: string) => {
      const b = bookmarks.find((x) => x.name === name);
      if (b) set({ ...view, ...readView(b.view), widths: view.widths });
    },
    save: (name: string) => schema && void patch("schemas", schema.id, { views: [...bookmarks.filter((b) => b.name !== name), { name, view: { search: view.search, filters: view.filters, sorts: view.sorts } }] }),
    remove: (name: string) => schema && void patch("schemas", schema.id, { views: bookmarks.filter((b) => b.name !== name) }),
  };
}

/* ----------------------------------- Page ----------------------------------- */

/**
 * Excel: a schema's Excel table. Data: its rows as a Notion-like table (every column, filters, sorts,
 * bookmarks); Schema: its columns (types, choices with colors, rules), written to the Excel file and to
 * the form with the user's buttons; Context: the notes given to the AI with it.
 */
function Excel() {
  const { schema: schemaId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/excel" });
  useLast("excel:schema", schemaId, (schema) => navigate({ search: { schema }, replace: true }));
  const schemas = useCollection("schemas");
  const schema = schemas.find((s) => s.id === schemaId);
  const { table, error, reload, write } = useTable(schema);
  const views = useView(schema);
  const { view } = views;
  const [open, setOpen] = useState<number>();
  /** Rows checked by the user (Excel indexes): the AI then works on them only. */
  const [checked, setChecked] = useState<number[]>([]);
  const [adding, setAdding] = useState(false);
  useEffect(() => (setChecked([]), setOpen(undefined)), [schemaId]);
  const shown = useMemo(() => (table ? apply(table.rows, view, table.fields) : []), [table, view]);
  const excel = schema?.excel;

  const removeRows = (indexes: number[]) => {
    quiet(write(() => deleteRows(excel!, indexes), { optimistic: (t) => withoutRows(t, indexes) }));
    setChecked([]);
    setOpen(undefined);
  };
  /** New row: added at the end of the table (next id, default values), then opened to be filled. */
  const createRow = () => {
    if (!table) return;
    const index = table.rows.length;
    quiet(write((t) => addRows(excel!, t, [{}], { partial: true }), { refetch: true, optimistic: (t) => withRows(t, [{}]) }));
    setOpen(index);
  };
  // Uploads only when the schema has an uploads folder and the option is on.
  const folder = schema?.filesFolder;
  const upload: Uploader | undefined = EXCEL_UPLOAD && folder ? { save: (file) => uploadFileTo(folder, file), folder: folder.name } : undefined;

  const agent: Agent | undefined = useMemo(() => {
    if (!schema) return undefined;
    const structure = schema.builtIn ? undefined : schemaAgent(schema);
    const data = table && excel ? excelData({ schema, table, shown: () => shown.map((s) => s.index), checked, view, setView: views.update, openRow: setOpen, write }) : undefined;
    if (!data && !structure) return undefined;
    return {
      instructions: [
        `Screen: EXCEL. One Excel table on SharePoint ("${schema.title}") with two parts: its DATA (rows) and its SCHEMA (the columns' definition). Do what the request is about: rows → data tools; columns, choices, rules, form → schema tools.`,
        data?.instructions,
        structure?.instructions ?? (schema.builtIn ? "This is a built-in schema: its columns are fixed (no schema change)." : ""),
      ]
        .filter(Boolean)
        .join("\n\n"),
      tools: [...(data?.tools ?? []), ...(structure?.tools ?? [])],
      context: (data ?? structure)!.context,
      state: async () => [await data?.state?.(), await structure?.state?.()].filter(Boolean).join("\n\n"),
      finish: structure?.finish,
    };
  }, [schema, table, shown, checked, view, write]); // eslint-disable-line react-hooks/exhaustive-deps

  const rename = () => {
    if (!schema) return;
    const title = prompt("Name of this table in the app", schema.title)?.trim();
    if (title) void patch("schemas", schema.id, { title });
  };

  return (
    <Workspace
      title="Excel"
      icon={FileSpreadsheetIcon}
      picker={
        <>
          <EntityPicker
            items={schemas.map((s) => ({ id: s.id, label: s.title, hint: s.builtIn ? "built-in" : s.form ? "form" : undefined }))}
            value={schemaId}
            onChange={(schema) => navigate({ search: { schema } })}
            placeholder={schemas.length ? "Choose an Excel table" : "Add an Excel table"}
            onAdd={() => setAdding(true)}
            addLabel="Add an Excel table…"
          />
          <AddDialog
            open={adding}
            onOpenChange={setAdding}
            title="Add an Excel table"
            description="Paste the link of an Excel file on SharePoint or OneDrive. Its columns are detected (types from the data, or from its _schema sheet); a form link gives the questions' types and choices."
            fields={[
              { name: "excel", label: "Excel file link", placeholder: "https://…sharepoint.com/…/file.xlsx" },
              { name: "table", label: "Table name", optional: true, hint: "Default: the file's first table." },
              { name: "form", label: "Microsoft Form link", optional: true, hint: "The form whose answers go to this table." },
              { name: "title", label: "Name in the app", optional: true },
            ]}
            onSubmit={async (v) => navigate({ search: { schema: (await addSchema({ excel: v.excel, table: v.table || undefined, form: v.form || undefined, title: v.title || undefined })).id } })}
          />
        </>
      }
      actions={
        schema && (
          <>
            {excel && (
              <>
                <Button variant="ghost" size="icon" className="size-9" aria-label="Read the file again" title="Read the file again" onClick={() => void reload()}>
                  <RefreshCwIcon />
                </Button>
                <Button variant="outline" size="sm" asChild>
                  <a href={excel.webUrl ?? excel.link} target="_blank" rel="noreferrer">
                    <ExternalLinkIcon /> Excel
                  </a>
                </Button>
              </>
            )}
            {schema.form && (
              <Button variant="outline" size="sm" asChild>
                <a href={designUrl(schema.form)} target="_blank" rel="noreferrer">
                  <ClipboardListIcon /> Form
                </a>
              </Button>
            )}
            <EntityMenu
              kind="schemas"
              entry={schema}
              onRename={schema.builtIn ? undefined : rename}
              onRemove={schema.builtIn ? undefined : () => confirm(`Remove "${schema.title}" from the app? (the Excel file and the form are kept)`) && void remove("schemas", schema.id).then(() => navigate({ search: {} }))}
            >
              <DropdownMenuItem onSelect={() => downloadText(`${schema.title}.schema.json`, JSON.stringify(schema, null, 2))}>
                <FileIcon /> Download the schema
              </DropdownMenuItem>
            </EntityMenu>
          </>
        )
      }
      ai={{ scope: `excel:${schema?.id}`, agent, placeholder: "Ask to find, fix or add rows, or to change the columns…", empty: "Choose an Excel table, or add one from its link." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        {
          id: "data",
          label: "Data",
          icon: Table2Icon,
          count: table?.rows.length,
          content: !schema ? (
            <Empty icon={FileSpreadsheetIcon}>Choose an Excel table, or add one from its link (list at the top).</Empty>
          ) : !excel ? (
            <Empty icon={FileSpreadsheetIcon}>This built-in table has no Excel file yet: choose one in the Schema tab.</Empty>
          ) : (
            <>
              {error && <p className="border-b bg-destructive/5 px-6 py-2 text-[13px] text-destructive">{error}</p>}
              {!table && !error && <Empty>Loading…</Empty>}
              {table && (
                <>
                  <DataBar table={table} views={views} count={shown.length} checked={checked} onClear={() => setChecked([])} onDelete={() => confirm(`Delete ${checked.length} row(s) from the Excel file?`) && removeRows(checked)} onNew={createRow} />
                  <DataTableView schema={schema} table={table} items={shown} view={view} setView={views.update} upload={upload} write={write} checked={checked} setChecked={setChecked} onOpen={setOpen} />
                  <RowSheet schema={schema} table={table} index={open} onClose={() => setOpen(undefined)} write={write} upload={upload} onDelete={(i) => removeRows([i])} />
                </>
              )}
            </>
          ),
        },
        { id: "schema", label: "Schema", icon: Settings2Icon, count: schema?.fields.length, content: schema ? <SchemaTab key={schema.id} schema={schema} onWritten={() => void reload()} /> : <Empty>Choose an Excel table.</Empty> },
        { id: "context", label: "Context", icon: BookTextIcon, count: schema?.contextIds.length, content: schema ? <ContextLinks what="Excel table" ids={schema.contextIds} onChange={(contextIds) => void patch("schemas", schema.id, { contextIds })} /> : <Empty>Choose an Excel table.</Empty> },
      ]}
    />
  );
}

/* --------------------------------- Data bar --------------------------------- */

const chip = "h-8 gap-1.5 rounded-full px-3 text-[13px]";

/** Above the rows: bookmarks, filters, sorts, search, the selection and "New row". */
function DataBar({ table, views, count, checked, onClear, onDelete, onNew }: { table: Table; views: ReturnType<typeof useView>; count: number; checked: number[]; onClear: () => void; onDelete: () => void; onNew: () => void }) {
  const { view, update: setView } = views;
  const fields = table.fields;
  return (
    <div className="flex shrink-0 flex-col gap-2 border-b bg-card px-5 py-2.5">
      <div className="flex flex-wrap items-center gap-1.5">
        <BookmarkIcon className="mr-1 size-4 text-muted-foreground" aria-hidden />
        {views.bookmarks.map((b) => (
          <span key={b.name} className={cn("group flex h-7 items-center gap-1 rounded-full border pr-1 pl-3 text-[12.5px] transition-colors", b.active ? "border-primary/40 bg-primary/10 text-primary" : "hover:bg-muted")}>
            <button type="button" className="cursor-pointer" onClick={() => views.open(b.name)}>
              {b.name}
            </button>
            <button type="button" className="flex size-5 cursor-pointer items-center justify-center rounded-full opacity-0 group-hover:opacity-70 hover:bg-muted" aria-label={`Delete the bookmark ${b.name}`} onClick={() => confirm(`Delete the bookmark "${b.name}"?`) && views.remove(b.name)}>
              <XIcon className="size-3" />
            </button>
          </span>
        ))}
        {!views.bookmarks.length && <span className="text-[12.5px] text-muted-foreground">No bookmark yet: filter the rows, then save them.</span>}
        <Button
          variant="ghost"
          size="sm"
          className="h-7 rounded-full px-2.5 text-[12.5px] text-primary"
          onClick={() => {
            const name = prompt("Name of the bookmark (keeps the search, filters and sorts)")?.trim();
            if (name) views.save(name);
          }}
        >
          <BookmarkPlusIcon /> Save as bookmark
        </Button>
      </div>
      <div className="flex flex-wrap items-center gap-2">
        <Popover>
          <PopoverTrigger asChild>
            <Button variant="outline" size="sm" className={chip}>
              <SlidersHorizontalIcon /> Filter
              {view.filters.length > 0 && <span className="flex size-5 items-center justify-center rounded-full bg-primary text-[11px] text-primary-foreground">{view.filters.length}</span>}
            </Button>
          </PopoverTrigger>
          <PopoverContent align="start" className="w-[360px]">
            <FiltersEditor fields={fields} view={view} setView={setView} />
          </PopoverContent>
        </Popover>
        <Popover>
          <PopoverTrigger asChild>
            <Button variant="outline" size="sm" className={chip}>
              <ArrowUpDownIcon /> Sort
              {view.sorts.length > 0 && <span className="flex size-5 items-center justify-center rounded-full bg-primary text-[11px] text-primary-foreground">{view.sorts.length}</span>}
            </Button>
          </PopoverTrigger>
          <PopoverContent align="start" className="w-[300px]">
            <SortsEditor fields={fields} view={view} setView={setView} />
          </PopoverContent>
        </Popover>
        {fields
          .filter((f) => f.type === "option")
          .slice(0, 3)
          .map((f) => {
            const current = view.filters.find((x) => x.field === f.name && x.op === "is");
            const others = view.filters.filter((x) => x !== current);
            return (
              <Popover key={f.name}>
                <PopoverTrigger asChild>
                  <Button variant="outline" size="sm" className={cn(chip, current?.values?.length && "border-primary/40 bg-primary/5")}>
                    {labelOf(f)}
                    {current?.values?.length ? <span className="max-w-40 truncate text-muted-foreground">: {current.values.join(", ")}</span> : null}
                    <ChevronDownIcon />
                  </Button>
                </PopoverTrigger>
                <PopoverContent align="start" className="w-60 p-1.5">
                  {f.options?.map((o) => {
                    const on = !!current?.values?.includes(o.value);
                    const values = on ? current!.values!.filter((v) => v !== o.value) : [...(current?.values ?? []), o.value];
                    return (
                      <button key={o.value} type="button" className="flex w-full cursor-pointer items-center gap-2 rounded-md px-2 py-1.5 hover:bg-muted" onClick={() => setView({ filters: values.length ? [...others, { field: f.name, op: "is", values }] : others })}>
                        <Checkbox checked={on} tabIndex={-1} />
                        <Badge color={o.color}>{o.value}</Badge>
                      </button>
                    );
                  })}
                </PopoverContent>
              </Popover>
            );
          })}
        {(view.filters.length > 0 || view.sorts.length > 0 || view.search) && (
          <Button variant="ghost" size="sm" className="h-8 text-[12.5px] text-muted-foreground" onClick={() => setView({ filters: [], sorts: [], search: "" })}>
            Clear
          </Button>
        )}
        <span className="text-[13px] font-medium tabular-nums">
          {count} <span className="font-normal text-muted-foreground">of {table.rows.length} rows</span>
        </span>
        {checked.length > 0 && (
          <span className="flex h-8 items-center gap-1 rounded-full bg-foreground pr-1 pl-3 text-[12.5px] text-background">
            {checked.length} selected
            <button type="button" className="flex size-6 cursor-pointer items-center justify-center rounded-full hover:bg-white/15" aria-label="Delete the selected rows" onClick={onDelete}>
              <Trash2Icon className="size-3.5" />
            </button>
            <button type="button" className="flex size-6 cursor-pointer items-center justify-center rounded-full hover:bg-white/15" aria-label="Clear the selection" onClick={onClear}>
              <XIcon className="size-3.5" />
            </button>
          </span>
        )}
        <span className="ml-auto" />
        <label className="flex h-8 w-56 cursor-text items-center gap-2 rounded-full border bg-background px-3 text-[13px] focus-within:ring-2 focus-within:ring-ring/40">
          <SearchIcon className="size-4 text-muted-foreground" />
          <input className="min-w-0 flex-1 bg-transparent outline-none placeholder:text-muted-foreground" placeholder="Search" value={view.search ?? ""} onChange={(e) => setView({ search: e.target.value })} />
        </label>
        <Button size="sm" className="h-8 rounded-full" onClick={onNew}>
          <PlusIcon /> New row
        </Button>
      </div>
    </div>
  );
}

/** Every filter of the view: field, operator suited to its type, value(s). */
function FiltersEditor({ fields, view, setView }: { fields: Field[]; view: TableView; setView: (v: Partial<TableView>) => void }) {
  const set = (i: number, f?: Filter) => setView({ filters: f ? view.filters.map((x, k) => (k === i ? f : x)) : view.filters.filter((_, k) => k !== i) });
  return (
    <div className="flex flex-col gap-3">
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
              <Button variant="ghost" size="icon" className="size-8" aria-label="Remove the filter" onClick={() => set(i)}>
                <XIcon />
              </Button>
            </div>
            {!NO_VALUE.includes(filter.op) &&
              (field.type === "option" ? (
                <div className="flex flex-wrap gap-1">
                  {field.options?.map((o) => {
                    const on = filter.values?.includes(o.value);
                    return (
                      <button key={o.value} type="button" className={cn("cursor-pointer", on ? "" : "opacity-45 hover:opacity-100")} onClick={() => set(i, { ...filter, values: on ? filter.values!.filter((v) => v !== o.value) : [...(filter.values ?? []), o.value] })}>
                        <Badge color={o.color}>{o.value}</Badge>
                      </button>
                    );
                  })}
                </div>
              ) : (
                <Input className="h-8" type={field.type === "number" || field.type === "rating" ? "number" : field.type === "date" ? "date" : "text"} placeholder="Value" value={filter.value ?? ""} onChange={(e) => set(i, { ...filter, value: e.target.value })} />
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
    <div className="flex flex-col gap-2">
      {view.sorts.map((s, i) => {
        const field = fields.find((f) => f.name === s.field);
        return (
          <div key={i} className="flex items-center gap-1.5">
            <span className="flex-1 truncate text-[13px]">{field ? labelOf(field) : s.field}</span>
            <Button variant="outline" size="sm" className="h-8" onClick={() => setView({ sorts: view.sorts.map((x, k) => (k === i ? { ...x, desc: !x.desc } : x)) })}>
              {s.desc ? <ArrowDownIcon /> : <ArrowUpIcon />} {s.desc ? "Z → A" : "A → Z"}
            </Button>
            <Button variant="ghost" size="icon" className="size-8" aria-label="Remove the sort" onClick={() => setView({ sorts: view.sorts.filter((_, k) => k !== i) })}>
              <XIcon />
            </Button>
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

const defaultWidth = (f: Field) => (isIdField(f) ? 130 : f.type === "text" ? 320 : f.type === "string" ? 220 : f.type === "file" || f.type === "url" ? 200 : f.type === "image" ? 110 : 170);

/** Every column, edited in place (as in Notion); a row opens on the side with its expand button. */
function DataTableView({ schema, table, items, view, setView, upload, write, checked, setChecked, onOpen }: {
  schema: SchemaEntry;
  table: Table;
  items: { row: Row; index: number }[];
  view: TableView;
  setView: (v: Partial<TableView>) => void;
  upload?: Uploader;
  write: Write;
  checked: number[];
  setChecked: (c: number[]) => void;
  onOpen: (index: number) => void;
}) {
  const columns = table.fields;
  const all = items.length > 0 && items.every((i) => checked.includes(i.index));
  const change = (index: number, name: string, value: unknown) => {
    const changes = [{ index, values: { [name]: value } }];
    quiet(write((t) => updateRows(schema.excel!, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  // Drag a header's right edge to resize the column (kept in this browser when released).
  const [resizing, setResizing] = useState<{ name: string; width: number }>();
  const widthOf = (f: Field) => (resizing?.name === f.name ? resizing.width : view.widths?.[f.name]) ?? defaultWidth(f);
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
    <div className="min-h-0 flex-1 overflow-auto bg-card">
      <table className="w-max min-w-full border-separate border-spacing-0 text-[13.5px]">
        <thead className="sticky top-0 z-10 bg-muted/70 backdrop-blur">
          <tr>
            <th className="sticky left-0 z-20 w-20 border-r border-b bg-muted">
              <span className="flex justify-center">
                <Checkbox aria-label="Select all the rows shown" checked={all} onCheckedChange={() => setChecked(all ? [] : items.map((i) => i.index))} />
              </span>
            </th>
            {columns.map((f) => {
              const sort = view.sorts.find((s) => s.field === f.name);
              return (
                <th key={f.name} style={style(f)} className="relative border-r border-b p-0 last:border-r-0">
                  <button type="button" title={f.description ?? "Sort"} onClick={() => sortBy(f)} className="th flex w-full cursor-pointer items-center gap-2 hover:bg-muted [&_svg]:size-3.5 [&_svg]:shrink-0 [&_svg]:text-muted-foreground">
                    {TYPE_ICONS[f.type]}
                    <span className="truncate">{labelOf(f)}</span>
                    {f.required && <span className="text-destructive/70">*</span>}
                    {sort && (sort.desc ? <ArrowDownIcon /> : <ArrowUpIcon />)}
                  </button>
                  <div role="separator" aria-orientation="vertical" title="Drag to resize" onPointerDown={(e) => resize(f, e)} className={cn("absolute inset-y-0 -right-1 z-10 w-2 cursor-col-resize hover:bg-primary/25", resizing?.name === f.name && "bg-primary/35")} />
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {items.map(({ row, index }) => (
            <tr key={index} className={cn("group", checked.includes(index) ? "bg-primary/5" : "hover:bg-muted/40")}>
              <td className="sticky left-0 z-[5] border-r border-b bg-card align-middle group-hover:bg-muted">
                <span className="flex items-center justify-center gap-1">
                  <Checkbox aria-label="Select this row" checked={checked.includes(index)} onCheckedChange={() => setChecked(checked.includes(index) ? checked.filter((i) => i !== index) : [...checked, index])} />
                  <button type="button" className="flex size-6 cursor-pointer items-center justify-center rounded text-muted-foreground opacity-0 group-hover:opacity-100 hover:bg-background hover:text-foreground" aria-label="Open the row" title="Open the row" onClick={() => onOpen(index)}>
                    <Maximize2Icon className="size-3.5" />
                  </button>
                </span>
              </td>
              {columns.map((f) => (
                <td key={f.name} style={style(f)} className="overflow-hidden border-r border-b p-0 align-middle last:border-r-0">
                  <FieldInput variant="cell" field={f} value={row[f.name]} upload={upload} validate={(v) => idError(table, index, v)} onChange={(v) => change(index, f.name, v)} />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {items.length === 0 && <p className="p-10 text-center text-muted-foreground">No row matches the filters.</p>}
    </div>
  );
}

/** One row with every field as a form, on the side; changes are written at once. */
function RowSheet({ schema, table, index, onClose, write, upload, onDelete }: { schema: SchemaEntry; table: Table; index?: number; onClose: () => void; write: Write; upload?: Uploader; onDelete: (index: number) => void }) {
  const row = index === undefined ? undefined : table.rows[index];
  const id = table.fields.find(isIdField);
  const set = (name: string, value: unknown) => {
    const changes = [{ index: index!, values: { [name]: value } }];
    quiet(write((t) => updateRows(schema.excel!, t, changes), { optimistic: (t) => withUpdates(t, changes) }));
  };
  return (
    <Sheet open={index !== undefined} onOpenChange={(o) => !o && onClose()}>
      <SheetContent className="w-full gap-0 sm:max-w-xl">
        <SheetHeader className="border-b">
          <SheetTitle>{row && id ? String(row[id.name] ?? "") || "New row" : "Row"}</SheetTitle>
          <SheetDescription>{schema.title} · row {index === undefined ? "" : index + 1}</SheetDescription>
        </SheetHeader>
        {row ? (
          <div className="flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto p-5">
            {table.fields.map((f) => (
              <div key={f.name} className="flex flex-col gap-1.5">
                <span className="flex items-center gap-1.5 text-[12.5px] font-medium text-muted-foreground [&_svg]:size-3.5">
                  {TYPE_ICONS[f.type]} {labelOf(f)}
                  {f.required && <span className="text-destructive/70">*</span>}
                </span>
                <FieldInput variant="form" field={f} value={row[f.name]} upload={upload} validate={(v) => idError(table, index!, v)} onChange={(v) => set(f.name, v)} />
                {f.description && <span className="text-[12px] text-muted-foreground">{f.description}</span>}
              </div>
            ))}
            <Button variant="outline" className="mt-2 ml-auto text-destructive" onClick={() => confirm("Delete this row from the Excel file?") && onDelete(index!)}>
              <Trash2Icon /> Delete the row
            </Button>
          </div>
        ) : (
          <Empty>Adding the row…</Empty>
        )}
      </SheetContent>
    </Sheet>
  );
}

/* ---------------------------------- Schema ---------------------------------- */

type Draft = { fields: Field[]; sections: FormSection[] };
type Busy = { busy?: string; done?: string; error?: string };

/**
 * The columns as a form: the list on the left, the chosen column's settings on the right, the form's
 * sections under it. Saved in the app with "Save"; written to the Excel file ("Sync the Excel":
 * columns, dropdowns, colors) and to the form ("Update the form", after review) by the user.
 */
function SchemaTab({ schema, onWritten }: { schema: SchemaEntry; onWritten: () => void }) {
  const saved: Draft = { fields: schema.fields, sections: schema.sections ?? [] };
  const [draft, setDraft] = useState<Draft>(saved);
  const [selected, setSelected] = useState<string | undefined>(schema.fields.find((f) => !isIdField(f))?.name ?? schema.fields[0]?.name);
  const [state, setState] = useState<Busy>({});
  const [plan, setPlan] = useState<FormPlan>();
  const savedText = JSON.stringify(saved);
  const changed = JSON.stringify(draft) !== savedText;
  // The AI (or an import) changed the schema: shown at once when nothing is being edited.
  const lastSaved = useRef(savedText);
  useEffect(() => {
    if (JSON.stringify(draft) === lastSaved.current) setDraft(saved);
    lastSaved.current = savedText;
  }, [savedText]); // eslint-disable-line react-hooks/exhaustive-deps
  const problems = useMemo(() => fieldsErrors(draft.fields, draft.sections), [draft]);
  const run = (busy: string, task: (progress: (label: string, done: number, total: number) => void) => Promise<string | void>) => {
    setState({ busy });
    task((label, done, total) => setState({ busy: `${label} ${done}/${total}` })).then((done) => setState({ done: done ?? undefined }), (e) => setState({ error: errorText(e) }));
  };
  const fixed = !!schema.builtIn;
  const field = draft.fields.find((f) => f.name === selected);
  const setField = (name: string, next: Field) => {
    setDraft({
      fields: draft.fields.map((f) => (f.name === name ? next : f)),
      // A renamed field keeps its place in its section.
      sections: next.name === name ? draft.sections : draft.sections.map((s) => ({ ...s, fields: s.fields.map((n) => (n === name ? next.name : n)) })),
    });
    if (next.name !== name) setSelected(next.name);
  };
  const move = (name: string, by: number) => {
    const i = draft.fields.findIndex((f) => f.name === name);
    const j = i + by;
    if (j < 0 || j >= draft.fields.length) return;
    const fields = [...draft.fields];
    [fields[i], fields[j]] = [fields[j], fields[i]];
    setDraft({ ...draft, fields });
  };
  const addField = () => {
    let name = "New column";
    for (let n = 2; draft.fields.some((f) => f.name === name); n++) name = `New column ${n}`;
    setDraft({ ...draft, fields: [...draft.fields, { name, type: "string" }] });
    setSelected(name);
  };
  const save = () => run("Saving…", async () => {
      await saveFields(schema, draft.fields, draft.sections);
      return "Saved. Use “Sync the Excel” to write it to the file.";
    });
  const syncExcel = () =>
    confirm(`Write the schema to "${schema.excel!.name}"?\n\nMissing columns are added${schema.form ? " at the end (the columns the form writes to never move)" : ""} — none is deleted —, the hidden _schema sheet is rewritten, and the table's sheet gets the dropdowns and colors of the choices.`) &&
    run("Writing the Excel file…", async (progress) => {
      await applySchemaToExcel(schema.excel!, schema.fields, progress, { appendOnly: !!schema.form });
      onWritten();
      return "Excel file updated.";
    });
  const reviewForm = () => run("Reading the form…", async () => void setPlan(await planForm(schema.form!, schema.fields, schema.sections)));

  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-y-auto">
      <div className="mx-auto flex w-full max-w-6xl flex-col gap-5 p-6">
        <Section
          title="Files"
          description={fixed ? "A built-in table: its columns are fixed; one Excel file holds the rows of every report." : "Where the rows live and what fills them. Saving the columns changes the app only; write them to the files with these buttons."}
          actions={
            !fixed &&
            schema.excel && (
              <>
                <StatusText state={state} />
                <Button variant="outline" size="sm" disabled={!schema.form || !!state.busy || changed} title={schema.form ? "Write the questions of the form after the schema (review first)" : "No form linked"} onClick={reviewForm}>
                  <UploadIcon /> Update the form
                </Button>
                <Button size="sm" disabled={!!state.busy || changed} title={changed ? "Save the columns first" : "Write the columns, dropdowns and colors to the Excel file"} onClick={syncExcel}>
                  <FileSpreadsheetIcon /> Sync the Excel
                </Button>
              </>
            )
          }
        >
          <div className="grid gap-4 md:grid-cols-3">
            <FileInfo label="Excel table">
              {schema.excel ? (
                <a className="inline-flex items-center gap-1 text-primary hover:underline" href={schema.excel.webUrl ?? schema.excel.link} target="_blank" rel="noreferrer">
                  {schema.excel.name} › {schema.excel.table} <ExternalLinkIcon className="size-3.5" />
                </a>
              ) : (
                <span className="text-muted-foreground">None</span>
              )}
              {fixed && <BuiltInExcel schema={schema} />}
            </FileInfo>
            {!fixed && (
              <FileInfo label="Microsoft Form">
                <FormLink schema={schema} />
              </FileInfo>
            )}
            {EXCEL_UPLOAD && !fixed && (
              <FileInfo label="Uploads folder">
                <UploadsFolder schema={schema} />
              </FileInfo>
            )}
          </div>
        </Section>
        {plan && (
          <FormReview
            plan={plan}
            onCancel={() => setPlan(undefined)}
            onApply={() => {
              const current = plan;
              setPlan(undefined);
              run("Writing the form…", async (progress) => {
                const { fields, sections } = await applyForm(schema.form!, current, progress);
                await patch("schemas", schema.id, { fields, sections: sections.length ? sections : undefined });
                return `Form updated: ${current.create.length} added, ${current.update.length} changed, ${current.remove.length} deleted.`;
              });
            }}
          />
        )}
        <Section
          title="Columns"
          description="Each column of the Excel table (and question of the form, named after the column)."
          actions={
            !fixed && (
              <>
                {changed && (
                  <Button variant="ghost" size="sm" onClick={() => (setDraft(saved), setState({}))}>
                    Discard
                  </Button>
                )}
                <Button size="sm" variant={changed ? "default" : "outline"} disabled={!changed || problems.length > 0 || !!state.busy} onClick={save}>
                  <SaveIcon /> Save
                </Button>
              </>
            )
          }
        >
          {problems.length > 0 && <p className="mb-4 rounded-lg bg-destructive/10 px-3 py-2 text-[13px] text-destructive">{problems.join(" ")}</p>}
          <div className="grid gap-5 md:grid-cols-[280px_1fr]">
            <div className="flex flex-col gap-1">
              {draft.fields.map((f) => (
                <button
                  key={f.name}
                  type="button"
                  onClick={() => setSelected(f.name)}
                  className={cn("flex cursor-pointer items-center gap-2 rounded-lg px-3 py-2 text-left text-[13.5px] transition-colors [&_svg]:size-4 [&_svg]:shrink-0", f.name === selected ? "bg-primary/10 text-primary" : "hover:bg-muted")}
                >
                  <span className={f.name === selected ? "" : "text-muted-foreground"}>{TYPE_ICONS[f.type]}</span>
                  <span className="min-w-0 flex-1 truncate">{labelOf(f)}</span>
                  {f.required && <span className="text-destructive/70">*</span>}
                  {f.form === false && <span className="text-[11px] text-muted-foreground">Excel only</span>}
                </button>
              ))}
              {!fixed && (
                <Button variant="ghost" size="sm" className="mt-1 justify-start text-primary" onClick={addField}>
                  <PlusIcon /> Add a column
                </Button>
              )}
            </div>
            {field ? (
              <FieldEditor
                key={draft.fields.indexOf(field)}
                field={field}
                fields={draft.fields}
                sections={draft.sections}
                readOnly={fixed || isIdField(field)}
                onChange={(next) => setField(field.name, next)}
                onMove={(by) => move(field.name, by)}
                onDelete={() => {
                  setDraft({ fields: draft.fields.filter((f) => f !== field), sections: draft.sections.map((s) => ({ ...s, fields: s.fields.filter((n) => n !== field.name) })) });
                  setSelected(draft.fields.find((f) => f !== field)?.name);
                }}
              />
            ) : (
              <Empty>Choose a column.</Empty>
            )}
          </div>
        </Section>
        {!fixed && schema.form && <SectionsEditor draft={draft} onChange={(sections) => setDraft({ ...draft, sections })} />}
      </div>
    </div>
  );
}

function FileInfo({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="flex min-w-0 flex-col gap-1.5">
      <span className="text-[12px] font-medium text-muted-foreground">{label}</span>
      <div className="flex min-w-0 flex-col gap-2 text-[13.5px]">{children}</div>
    </div>
  );
}

/** The settings of one column, by type. */
function FieldEditor({ field, fields, sections, readOnly, onChange, onMove, onDelete }: { field: Field; fields: Field[]; sections: FormSection[]; readOnly: boolean; onChange: (f: Field) => void; onMove: (by: number) => void; onDelete: () => void }) {
  const [name, setName] = useState(field.name);
  const set = (change: Partial<Field>) => {
    const next = { ...field, ...change };
    for (const k of Object.keys(change) as (keyof Field)[]) if (next[k] === undefined || next[k] === "") delete next[k];
    onChange(next);
  };
  const options = field.options ?? [];
  const setOption = (i: number, o?: Option) => set({ options: o ? options.map((x, k) => (k === i ? o : x)) : options.filter((_, k) => k !== i) });
  const targets = [...sections.map((s) => s.title), "end"];
  return (
    <div className="flex flex-col gap-4 rounded-xl border bg-background p-5">
      <div className="grid gap-4 sm:grid-cols-2">
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="field-name">Column name (= question)</Label>
          <Input id="field-name" value={name} disabled={readOnly} onChange={(e) => setName(e.target.value)} onBlur={() => name.trim() && name.trim() !== field.name && set({ name: name.trim() })} />
          {field.question && <p className="text-[11.5px] text-warning">Renaming a column linked to the form creates a new, empty column in the Excel.</p>}
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="field-type">Type</Label>
          <select id="field-type" className="input h-9" value={field.type} disabled={readOnly} onChange={(e) => set({ type: e.target.value as FieldType })}>
            {FIELD_TYPES.map((t) => (
              <option key={t} value={t}>{TYPE_LABELS[t]}</option>
            ))}
          </select>
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="field-label">Label in the app</Label>
          <Input id="field-label" placeholder={field.name} value={field.label ?? ""} disabled={readOnly} onChange={(e) => set({ label: e.target.value })} />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="field-default">Default value (new rows)</Label>
          <Input id="field-default" value={Array.isArray(field.default) ? field.default.join("; ") : (field.default ?? "").toString()} disabled={readOnly} onChange={(e) => set({ default: field.type === "number" || field.type === "rating" ? (e.target.value === "" ? undefined : Number(e.target.value)) : e.target.value })} />
        </div>
      </div>
      <div className="flex flex-col gap-1.5">
        <Label htmlFor="field-description">Description (help text, the question's subtitle)</Label>
        <Textarea id="field-description" rows={2} value={field.description ?? ""} disabled={readOnly} onChange={(e) => set({ description: e.target.value })} />
      </div>
      <div className="flex flex-wrap gap-x-6 gap-y-3">
        <Toggle label="Required" checked={!!field.required} disabled={readOnly} onChange={(v) => set({ required: v || undefined })} />
        <Toggle label="In the form" checked={field.form !== false} disabled={readOnly} onChange={(v) => set({ form: v ? undefined : false })} />
        {field.type === "option" && <Toggle label="Several values" checked={!!field.multiple} disabled={readOnly} onChange={(v) => set({ multiple: v || undefined, ...(v && { branches: undefined }) })} />}
        {field.type === "rating" && (
          <label className="flex items-center gap-2 text-[13px]">
            Highest
            <select className="input h-8 w-20" value={field.max ?? 5} disabled={readOnly} onChange={(e) => set({ max: Number(e.target.value) })}>
              <option value={5}>5</option>
              <option value={10}>10</option>
            </select>
          </label>
        )}
      </div>
      {field.type === "option" && (
        <div className="flex flex-col gap-2">
          <Label>Choices</Label>
          {options.map((o, i) => (
            <div key={i} className="flex items-center gap-2">
              <Popover>
                <PopoverTrigger asChild>
                  <button type="button" disabled={readOnly} className="size-7 shrink-0 cursor-pointer rounded-full border-2 border-background ring-1 ring-border" style={{ background: SWATCHES[o.color] }} aria-label={`Color of ${o.value}: ${o.color}`} />
                </PopoverTrigger>
                <PopoverContent className="flex w-auto gap-1.5 p-2">
                  {COLORS.map((c) => (
                    <button key={c} type="button" className={cn("size-6 cursor-pointer rounded-full ring-offset-2", c === o.color && "ring-2 ring-ring")} style={{ background: SWATCHES[c] }} aria-label={c} onClick={() => setOption(i, { ...o, color: c })} />
                  ))}
                </PopoverContent>
              </Popover>
              <Input className="h-8" value={o.value} disabled={readOnly} onChange={(e) => setOption(i, { ...o, value: e.target.value })} />
              {!field.multiple && sections.length > 0 && (
                <select className="input h-8 w-44" title="Where the form goes after this answer" value={field.branches?.[o.value] ?? ""} disabled={readOnly} onChange={(e) => {
                  const branches = { ...field.branches };
                  if (e.target.value) branches[o.value] = e.target.value;
                  else delete branches[o.value];
                  set({ branches: Object.keys(branches).length ? branches : undefined });
                }}>
                  <option value="">Then: next question</option>
                  {targets.map((t) => (
                    <option key={t} value={t}>{t === "end" ? "Then: end of the form" : `Then: ${t}`}</option>
                  ))}
                </select>
              )}
              <Button variant="ghost" size="icon" className="size-8" disabled={readOnly} aria-label={`Remove ${o.value}`} onClick={() => setOption(i)}>
                <XIcon />
              </Button>
              <Badge color={o.color}>{o.value || "…"}</Badge>
            </div>
          ))}
          {!readOnly && (
            <Button variant="ghost" size="sm" className="w-fit text-primary" onClick={() => set({ options: [...options, { value: `Choice ${options.length + 1}`, color: COLORS[(options.length + 1) % COLORS.length] }] })}>
              <PlusIcon /> Add a choice
            </Button>
          )}
        </div>
      )}
      {!readOnly && (
        <div className="flex items-center gap-2 border-t pt-4">
          <Button variant="outline" size="sm" disabled={fields[0] === field} onClick={() => onMove(-1)}>
            <ArrowUpIcon /> Up
          </Button>
          <Button variant="outline" size="sm" disabled={fields.at(-1) === field} onClick={() => onMove(1)}>
            <ArrowDownIcon /> Down
          </Button>
          <Button variant="ghost" size="sm" className="ml-auto text-destructive" onClick={() => confirm(`Remove the column "${field.name}" from the schema? (the Excel column and its values are kept)`) && onDelete()}>
            <Trash2Icon /> Remove
          </Button>
        </div>
      )}
    </div>
  );
}

function Toggle({ label, checked, disabled, onChange }: { label: string; checked: boolean; disabled?: boolean; onChange: (v: boolean) => void }) {
  return (
    <label className="flex cursor-pointer items-center gap-2 text-[13px]">
      <Switch checked={checked} disabled={disabled} onCheckedChange={onChange} />
      {label}
    </label>
  );
}

/** The form's sections: title, questions, where the form goes when done. */
function SectionsEditor({ draft, onChange }: { draft: Draft; onChange: (s: FormSection[]) => void }) {
  const { sections, fields } = draft;
  const set = (i: number, s?: FormSection) => onChange(s ? sections.map((x, k) => (k === i ? s : x)) : sections.filter((_, k) => k !== i));
  const inForm = fields.filter((f) => f.form !== false && !isIdField(f));
  const free = inForm.filter((f) => !sections.some((s) => s.fields.includes(f.name)));
  return (
    <Section
      title="Form sections"
      description={`Pages of the form. Questions in no section come first${free.length ? `: ${free.map(labelOf).join(", ")}` : ""}.`}
      actions={
        <Button variant="outline" size="sm" onClick={() => onChange([...sections, { title: `Section ${sections.length + 1}`, fields: [] }])}>
          <PlusIcon /> Add a section
        </Button>
      }
    >
      {!sections.length && <p className="text-[13px] text-muted-foreground">No section: the form is one page.</p>}
      <div className="flex flex-col gap-3">
        {sections.map((s, i) => (
          <div key={i} className="flex flex-col gap-3 rounded-xl border bg-background p-4">
            <div className="flex gap-2">
              <Input className="h-8 font-medium" value={s.title} onChange={(e) => set(i, { ...s, title: e.target.value })} />
              <select className="input h-8 w-56" value={s.next ?? ""} onChange={(e) => set(i, { ...s, next: e.target.value || undefined })}>
                <option value="">Then: next section</option>
                {sections.filter((x) => x !== s).map((x) => (
                  <option key={x.title} value={x.title}>Then: {x.title}</option>
                ))}
                <option value="end">Then: end of the form</option>
              </select>
              <Button variant="ghost" size="icon" className="size-8" aria-label="Remove the section" onClick={() => set(i)}>
                <XIcon />
              </Button>
            </div>
            <div className="flex flex-wrap gap-1.5">
              {inForm.map((f) => {
                const on = s.fields.includes(f.name);
                const elsewhere = !on && sections.some((x) => x !== s && x.fields.includes(f.name));
                if (elsewhere) return null;
                return (
                  <button key={f.name} type="button" className={cn("h-7 cursor-pointer rounded-full border px-3 text-[12.5px] transition-colors", on ? "border-primary/40 bg-primary/10 text-primary" : "text-muted-foreground hover:bg-muted")} onClick={() => set(i, { ...s, fields: on ? s.fields.filter((n) => n !== f.name) : [...s.fields, f.name] })}>
                    {labelOf(f)}
                  </button>
                );
              })}
            </div>
          </div>
        ))}
      </div>
    </Section>
  );
}

/** The Excel file of a built-in table (one for every report): a file without table gets one. */
function BuiltInExcel({ schema }: { schema: SchemaEntry }) {
  const [link, setLink] = useState("");
  return (
    <form
      className="flex gap-2"
      onSubmit={(e) => {
        e.preventDefault();
        if (link.trim()) void withToast("Setting up the Excel file…", setBuiltInExcel(schema, link.trim()), "Excel file set.").then(() => setLink(""));
      }}
    >
      <Input className="h-8" placeholder={schema.excel ? "Another Excel file link" : "Excel file link"} value={link} onChange={(e) => setLink(e.target.value)} />
      <Button type="submit" size="sm" variant="outline" disabled={!link.trim()}>
        Use
      </Button>
    </form>
  );
}

/** The form linked to the table: its name, or a field to link one (its questions can retype the matching columns). */
function FormLink({ schema }: { schema: SchemaEntry }) {
  const [link, setLink] = useState("");
  if (schema.form)
    return (
      <span className="flex items-center gap-1.5">
        <a href={designUrl(schema.form)} target="_blank" rel="noreferrer" className="truncate text-primary hover:underline">
          {schema.form.title || "Form"}
        </a>
        <Button variant="ghost" size="icon" className="size-7" aria-label="Unlink the form" title="Unlink the form (the form is kept)" onClick={() => confirm("Unlink the form from this table? (the form itself is kept)") && void patch("schemas", schema.id, { form: undefined })}>
          <XIcon />
        </Button>
      </span>
    );
  return (
    <form
      className="flex gap-2"
      onSubmit={(e) => {
        e.preventDefault();
        if (link.trim()) void withToast("Reading the form…", linkForm(schema, link.trim(), confirm("Give the matching columns the types and choices of the form's questions?")), "Form linked.").then(() => setLink(""));
      }}
    >
      <Input className="h-8" placeholder="Form link (optional)" value={link} onChange={(e) => setLink(e.target.value)} />
      <Button type="submit" size="sm" variant="outline" disabled={!link.trim()}>
        <LinkIcon /> Link
      </Button>
    </form>
  );
}

/** Optional: the SharePoint folder receiving the files uploaded in the Data tab (image and file columns). */
function UploadsFolder({ schema }: { schema: SchemaEntry }) {
  const [link, setLink] = useState("");
  if (schema.filesFolder)
    return (
      <span className="flex items-center gap-1.5">
        <a href={schema.filesFolder.link} target="_blank" rel="noreferrer" className="flex items-center gap-1.5 truncate text-primary hover:underline">
          <FolderIcon className="size-4" /> {schema.filesFolder.name}
        </a>
        <Button variant="ghost" size="icon" className="size-7" aria-label="No uploads (links only)" onClick={() => void patch("schemas", schema.id, { filesFolder: undefined })}>
          <XIcon />
        </Button>
      </span>
    );
  return (
    <form
      className="flex gap-2"
      onSubmit={(e) => {
        e.preventDefault();
        if (link.trim()) void withToast("Reading the folder…", resolveFolder(link.trim()).then((filesFolder) => patch("schemas", schema.id, { filesFolder })), "Uploads folder set.").then(() => setLink(""));
      }}
    >
      <Input className="h-8" placeholder="Folder link (optional)" title="SharePoint folder receiving the uploaded images and files" value={link} onChange={(e) => setLink(e.target.value)} />
      <Button type="submit" size="sm" variant="outline" disabled={!link.trim()}>
        Use
      </Button>
    </form>
  );
}

/** What updating the form will do, to confirm (deleting a question deletes its answers in the form). */
function FormReview({ plan, onApply, onCancel }: { plan: FormPlan; onApply: () => void; onCancel: () => void }) {
  const nothing = !plan.create.length && !plan.update.length && !plan.remove.length;
  const list = (title: string, items: string[], color: "green" | "blue" | "red") =>
    items.length > 0 && (
      <div className="flex flex-col gap-1.5">
        <span className="text-[13px] font-medium">{title}</span>
        <span className="flex flex-wrap gap-1.5">
          {items.map((t, i) => (
            <Badge key={i} color={color}>{t}</Badge>
          ))}
        </span>
      </div>
    );
  return (
    <div className="flex flex-col gap-3 rounded-xl border border-warning/40 bg-warning/10 p-4">
      <span className="font-semibold">{nothing ? "The form already matches the schema." : "Update the form?"}</span>
      {list("To add", plan.create, "green")}
      {list("To change (title, type, choices, branches, order)", plan.update, "blue")}
      {list("Questions to delete (their answers are deleted from the form)", plan.remove.map((q) => q.title), "red")}
      <div className="flex gap-2">
        {!nothing && (
          <Button size="sm" onClick={onApply}>
            Update the form
          </Button>
        )}
        <Button size="sm" variant="outline" onClick={onCancel}>
          {nothing ? "Close" : "Cancel"}
        </Button>
      </div>
    </div>
  );
}

