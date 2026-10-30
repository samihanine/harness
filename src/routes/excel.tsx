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
import { EXCEL_UPLOAD, SCREEN_LABELS } from "@/config";
import { excelData } from "@/features/excel/excel-tools";
import type { Agent } from "@/agent/loop";
import { addRows, deleteRows, formatTable, idError, resolveFolder, updateRows, uploadFileTo, withRows, withUpdates, withoutRows, type Table } from "@/services/excel";
import { addSchema, downloadText } from "@/lib/library";
import { useLast } from "@/lib/last";
import { errorText } from "@/services/ms";
import { COLORS, FIELD_TYPES, SWATCHES, TYPE_LABELS, fieldsErrors, isIdField, labelOf, type Field, type FieldType, type FormSection, type Option, type Row } from "@/features/excel/schema";
import { patch, remove, useCollection, type SchemaEntry } from "@/lib/store";
import { EMPTY_VIEW, NO_VALUE, OPS, OPS_OF, apply, filterLabel, readView, type Filter, type TableView } from "@/features/excel/table-view";
import { useTable, type Write } from "@/features/excel/use-table";
import { cn } from "@/lib/utils";
import { excelInstructions } from "@/prompts/excel";

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
    if (!schema || !table || !excel) return undefined;
    const data = excelData({ schema, table, shown: () => shown.map((s) => s.index), checked, view, setView: views.update, openRow: setOpen, write });
    return {
      ...data,
      instructions: excelInstructions(schema, data),
    };
  }, [schema, table, shown, checked, view, write]); // eslint-disable-line react-hooks/exhaustive-deps

  const rename = () => {
    if (!schema) return;
    const title = prompt("Name of this table in the app", schema.title)?.trim();
    if (title) void patch("schemas", schema.id, { title });
  };

  return (
    <Workspace
      title={SCREEN_LABELS.excel}
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
            description="Paste the link of an Excel file on SharePoint or OneDrive. Its columns are the table's own, read from the file."
            fields={[
              { name: "excel", label: "Excel file link", placeholder: "https://…sharepoint.com/…/file.xlsx" },
              { name: "table", label: "Table name", optional: true, hint: "Default: the file's first table." },
              { name: "title", label: "Name in the app", optional: true },
            ]}
            onSubmit={async (v) => navigate({ search: { schema: (await addSchema({ excel: v.excel, table: v.table || undefined, title: v.title || undefined })).id } })}
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
      ai={{ scope: `excel:${schema?.id}`, agent, placeholder: "Ask to find, fix or add rows…", empty: "Choose an Excel table, or add one from its link." }}
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
                  {/* Every row the same height: a long text or many choices are cut (the whole value: open the row). */}
                  <div className="flex h-11 items-center overflow-hidden [&>*]:w-full">
                    <FieldInput variant="cell" field={f} value={row[f.name]} upload={upload} validate={(v) => idError(table, index, v)} onChange={(v) => change(index, f.name, v)} />
                  </div>
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
/**
 * The table as the Excel file has it: its columns, read live from the header row (the app never adds,
 * types or moves them), and "Sync the formatting" (grey header, borders, column widths, wrapped text).
 */
function SchemaTab({ schema, onWritten }: { schema: SchemaEntry; onWritten: () => void }) {
  const [state, setState] = useState<Busy>({});
  const sync = () => {
    setState({ busy: "Formatting the table…" });
    formatTable(schema.excel!).then(
      () => (onWritten(), setState({ done: "Formatting applied." })),
      (e) => setState({ error: errorText(e) }),
    );
  };
  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-y-auto">
      <div className="mx-auto flex w-full max-w-4xl flex-col gap-5 p-6">
        <Section
          title="Files"
          description="Where the rows live. The columns are the Excel table's own: change them in Excel, the app reads them again."
          actions={
            schema.excel && (
              <>
                <StatusText state={state} />
                <Button size="sm" disabled={!!state.busy} title="Grey bold header, borders, column widths, wrapped text, rows fitted to their text" onClick={sync}>
                  <FileSpreadsheetIcon /> Sync the formatting
                </Button>
              </>
            )
          }
        >
          <div className="grid gap-4 md:grid-cols-2">
            <FileInfo label="Excel table">
              {schema.excel ? (
                <a className="inline-flex items-center gap-1 text-primary hover:underline" href={schema.excel.webUrl ?? schema.excel.link} target="_blank" rel="noreferrer">
                  {schema.excel.name} › {schema.excel.table} <ExternalLinkIcon className="size-3.5" />
                </a>
              ) : (
                <span className="text-muted-foreground">None</span>
              )}
            </FileInfo>
            {EXCEL_UPLOAD && (
              <FileInfo label="Uploads folder">
                <UploadsFolder schema={schema} />
              </FileInfo>
            )}
          </div>
        </Section>
        <Section title="Columns" description="In the order of the Excel table, as read from the file.">
          <div className="flex flex-wrap gap-1.5">
            {schema.fields.map((f) => (
              <span key={f.name} className="rounded-md border bg-muted/40 px-2 py-1 text-[12.5px]">
                {f.name}
              </span>
            ))}
            {!schema.fields.length && <span className="text-[13px] text-muted-foreground">Not read yet: open the Data tab.</span>}
          </div>
        </Section>
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


