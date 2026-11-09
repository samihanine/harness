/**
 * Excels: an Excel table on SharePoint used as a simple database, Notion-like. Views as tabs (table,
 * board, list), filter chips, inline editing, a row opened on the side as a form; the columns (4
 * types) edited in their own tab. Every change goes to the file at once; the file's own changes are
 * read again (sync.ts). The assistant changes rows by condition, without ids.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import { useNavigate, useSearch } from "@tanstack/react-router";
import {
  AlignLeftIcon,
  ArrowDownIcon,
  ArrowUpIcon,
  BookTextIcon,
  CalendarIcon,
  CheckIcon,
  ChevronDownIcon,
  CircleAlertIcon,
  CircleDotIcon,
  ColumnsIcon,
  CopyIcon,
  DatabaseIcon,
  ExternalLinkIcon,
  EyeOffIcon,
  GripVerticalIcon,
  KanbanIcon,
  ListIcon,
  Loader2Icon,
  Maximize2Icon,
  PaintbrushIcon,
  PencilIcon,
  PlusIcon,
  RefreshCwIcon,
  SearchIcon,
  Table2Icon,
  Trash2Icon,
  TypeIcon,
  XIcon,
  type LucideIcon,
} from "lucide-react";
import { Badge, COLORS, SWATCHES, type Color } from "@/components/badge";
import { ContextLinks } from "@/components/context-links";
import { AddDialog, EntityMenu, withToast } from "@/components/entity";
import { Button } from "@/components/ui/button";
import { DropdownMenu, DropdownMenuCheckboxItem, DropdownMenuContent, DropdownMenuItem, DropdownMenuLabel, DropdownMenuSeparator, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from "@/components/ui/sheet";
import { EntityPicker, Empty, Section, Workspace } from "@/components/workspace";
import { useLast } from "@/lib/last";
import { patch, put, remove, useCollection } from "@/lib/store";
import { cn } from "@/lib/utils";
import { resolveExcel } from "@/services/excel";
import { excel as feature } from "./module";
import { useExcelTable, type ExcelTable, type SyncStatus } from "./sync";
import { COLUMN_TYPES, DATE_RANGES, DEFAULT_VIEW, TYPE_LABELS, filterText, viewRows, type Column, type ColumnType, type DateRange, type Filter, type Layout, type Row, type TableEntry, type View } from "./table";
import { excelAgent } from "./tools";

const TYPE_ICONS: Record<ColumnType, LucideIcon> = { text: TypeIcon, longtext: AlignLeftIcon, date: CalendarIcon, option: CircleDotIcon };
const LAYOUT_ICONS: Record<Layout, LucideIcon> = { table: Table2Icon, board: KanbanIcon, list: ListIcon };
const LAYOUT_LABELS: Record<Layout, string> = { table: "Table", board: "Board", list: "List" };

/** Adds a table of a file (its link; the given table, else the first one). */
async function addTable({ link, table, title }: { link: string; table?: string; title?: string }) {
  const file = await resolveExcel(link);
  const name = table && file.tables.includes(table) ? table : file.tables[0];
  if (!name) throw new Error(`"${file.name}" has no Excel table: select the data in Excel and use Insert › Table.`);
  const id = crypto.randomUUID();
  return put("tables", {
    id,
    title: title?.trim() || (file.tables.length > 1 ? `${file.name} › ${name}` : file.name),
    excel: { link, driveId: file.driveId, itemId: file.itemId, name: file.name, table: name, fileUrl: file.fileUrl, webUrl: file.webUrl },
    columns: [],
    views: [DEFAULT_VIEW],
    contextIds: [],
  });
}

export default function Excel() {
  const { table: tableId, tab } = useSearch({ from: "/excel" });
  const navigate = useNavigate({ from: "/excel" });
  useLast("excel:table", tableId, (table) => navigate({ search: { table }, replace: true }));
  const tables = useCollection("tables");
  const entry = tables.find((t) => t.id === tableId);
  const table = useExcelTable(entry);
  const [adding, setAdding] = useState(false);
  const [viewId, setViewId] = useState<string>();
  const views = entry?.views.length ? entry.views : [DEFAULT_VIEW];
  const view = views.find((v) => v.id === viewId) ?? views[0];
  const setView = (v: View) => void table.setViews(views.map((x) => (x.id === v.id ? v : x)));
  const live = useRef({ entry, table, view });
  live.current = { entry, table, view };
  const agent = entry && excelAgent(() => live.current.entry!, () => live.current.table, () => live.current.view, (v) => setView(v));

  return (
    <Workspace
      title={feature.label}
      icon={DatabaseIcon}
      picker={<EntityPicker items={tables.map((t) => ({ id: t.id, label: t.title }))} value={tableId} onChange={(table) => navigate({ search: { table } })} placeholder={tables.length ? "Choose a table" : "Add an Excel table"} onAdd={() => setAdding(true)} addLabel="Add an Excel table…" />}
      actions={
        <>
          {entry && <SyncPill status={table.status} dirty={table.dirty} onSync={() => void withToast("Writing the columns and formatting to Excel…", table.sync(), "Synced with Excel.")} />}
          {entry?.excel.webUrl && (
            <Button size="sm" variant="outline" asChild>
              <a href={entry.excel.webUrl} target="_blank" rel="noreferrer">
                <ExternalLinkIcon /> Excel
              </a>
            </Button>
          )}
          <EntityMenu
            kind="tables"
            entry={entry}
            onRename={entry && (() => { const title = prompt("Name of the table in the app", entry.title)?.trim(); if (title) void patch("tables", entry.id, { title }); })}
            onRefresh={entry && (() => void table.reload())}
            onRemove={entry && (() => confirm(`Remove "${entry.title}" from the app? (the Excel file is kept)`) && void remove("tables", entry.id).then(() => navigate({ search: {} })))}
          />
          <AddDialog
            open={adding}
            onOpenChange={setAdding}
            title="Add an Excel table"
            description="Paste the sharing link of an .xlsx file on SharePoint or OneDrive. Its columns are read with their types guessed; change them in the Columns tab."
            fields={[
              { name: "link", label: "Excel file link", placeholder: "https://…sharepoint.com/…/file.xlsx" },
              { name: "table", label: "Table name", optional: true, hint: "The first table of the file when empty." },
              { name: "title", label: "Name in the app", optional: true },
            ]}
            onSubmit={async (v) => navigate({ search: { table: (await addTable({ link: v.link, table: v.table || undefined, title: v.title || undefined })).id } })}
          />
        </>
      }
      ai={{ scope: `excel:${entry?.id}`, agent, placeholder: "Ask to update rows, add a column, filter the view…", empty: "Choose an Excel table to work on it with the assistant." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      notices={table.status.state === "error" ? [{ id: "sync", level: "error", title: "Excel refused a change", body: table.status.message }] : []}
      tabs={[
        {
          id: "data",
          label: "Data",
          icon: Table2Icon,
          count: table.data?.rows.length,
          content: entry ? (
            table.data ? (
              <Database entry={entry} table={table} views={views} view={view} onView={setViewId} />
            ) : (
              <Empty icon={table.status.state === "error" ? CircleAlertIcon : Loader2Icon}>{table.status.state === "error" ? table.status.message : "Reading the table…"}</Empty>
            )
          ) : (
            <Empty icon={DatabaseIcon} action={<Button onClick={() => setAdding(true)}><PlusIcon /> Add an Excel table</Button>}>
              An Excel table on SharePoint, used as a simple database: views, filters and columns of 4 types, kept in sync with the file.
            </Empty>
          ),
        },
        { id: "columns", label: "Columns", icon: ColumnsIcon, count: entry?.columns.length, content: entry && table.data ? <ColumnsPanel entry={entry} table={table} /> : <Empty>Choose an Excel table.</Empty> },
        { id: "context", label: "Context", icon: BookTextIcon, count: entry?.contextIds.length, content: entry ? <ContextLinks what="Excel table" ids={entry.contextIds} onChange={(contextIds) => void patch("tables", entry.id, { contextIds })} /> : <Empty>Choose an Excel table.</Empty> },
      ]}
    />
  );
}

/* --------------------------------- Sync state --------------------------------- */

function SyncPill({ status, dirty, onSync }: { status: SyncStatus; dirty: boolean; onSync: () => void }) {
  const busy = status.state === "reading" || status.state === "writing" || status.state === "syncing";
  const label = busy ? { reading: "Reading…", writing: "Saving…", syncing: "Syncing columns…" }[status.state as "reading"] : status.state === "error" ? "Not saved" : dirty ? "Columns not in Excel yet" : status.message ?? "Synced";
  const tone = status.state === "error" ? "bg-red-500/10 text-destructive" : dirty ? "bg-amber-500/15 text-amber-800 dark:text-amber-300" : "bg-emerald-500/10 text-emerald-700 dark:text-emerald-300";
  return (
    <button type="button" onClick={onSync} title="Write the columns, their lists and colors, and the formatting to the Excel file" className={cn("flex h-8 items-center gap-1.5 rounded-full px-3 text-[12.5px] font-medium", tone)}>
      {busy ? <Loader2Icon className="size-3.5 animate-spin" /> : status.state === "error" ? <CircleAlertIcon className="size-3.5" /> : dirty ? <RefreshCwIcon className="size-3.5" /> : <CheckIcon className="size-3.5" />}
      {label}
      {!busy && <span className="opacity-60">· Sync</span>}
    </button>
  );
}

/* ---------------------------------- Database ---------------------------------- */

function Database({ entry, table, views, view, onView }: { entry: TableEntry; table: ExcelTable; views: View[]; view: View; onView: (id: string) => void }) {
  const [open, setOpen] = useState<number>();
  const data = table.data!;
  const shown = useMemo(() => viewRows(data.rows, view), [data.rows, view]);
  const columns = entry.columns.filter((c) => !view.hidden?.includes(c.name));
  const setView = (v: Partial<View>) => void table.setViews(views.map((x) => (x.id === view.id ? { ...view, ...v } : x)));
  const addView = (layout: Layout) => {
    const v: View = { id: crypto.randomUUID(), name: LAYOUT_LABELS[layout], layout, filters: [], group: layout === "board" ? entry.columns.find((c) => c.type === "option")?.name : undefined };
    void table.setViews([...views, v]).then(() => onView(v.id));
  };
  const save = (index: number, values: Row) => void withToast("Saving…", table.update([{ index, values }]));
  const newRow = async (values: Row = {}) => {
    const [at] = await table.add([values]);
    setOpen(at);
  };
  return (
    <div className="flex min-h-0 flex-1 flex-col">
      <div className="flex items-end gap-1 border-b px-4 pt-2">
        {views.map((v) => {
          const Icon = LAYOUT_ICONS[v.layout];
          const active = v.id === view.id;
          return (
            <div key={v.id} className={cn("group flex h-9 items-center border-b-2", active ? "border-foreground" : "border-transparent")}>
              <button type="button" onClick={() => onView(v.id)} className={cn("flex items-center gap-1.5 px-2.5 text-[13px]", active ? "font-medium" : "text-muted-foreground hover:text-foreground")}>
                <Icon className="size-3.5" /> {v.name}
              </button>
              {active && (
                <DropdownMenu>
                  <DropdownMenuTrigger className="icon-btn size-6 opacity-60 hover:opacity-100" aria-label="View options">
                    <ChevronDownIcon />
                  </DropdownMenuTrigger>
                  <DropdownMenuContent align="start" className="w-48">
                    <DropdownMenuItem onSelect={() => { const name = prompt("Name of the view", v.name)?.trim(); if (name) setView({ name }); }}>
                      <PencilIcon /> Rename
                    </DropdownMenuItem>
                    <DropdownMenuItem onSelect={() => { const copy = { ...v, id: crypto.randomUUID(), name: `${v.name} (copy)` }; void table.setViews([...views, copy]).then(() => onView(copy.id)); }}>
                      <CopyIcon /> Duplicate
                    </DropdownMenuItem>
                    <DropdownMenuSeparator />
                    <DropdownMenuLabel>Layout</DropdownMenuLabel>
                    {(Object.keys(LAYOUT_LABELS) as Layout[]).map((l) => (
                      <DropdownMenuCheckboxItem key={l} checked={v.layout === l} onCheckedChange={() => setView({ layout: l, group: l === "board" ? (v.group ?? entry.columns.find((c) => c.type === "option")?.name) : v.group })}>
                        {LAYOUT_LABELS[l]}
                      </DropdownMenuCheckboxItem>
                    ))}
                    {views.length > 1 && (
                      <>
                        <DropdownMenuSeparator />
                        <DropdownMenuItem className="text-destructive" onSelect={() => void table.setViews(views.filter((x) => x.id !== v.id)).then(() => onView(views.find((x) => x.id !== v.id)!.id))}>
                          <Trash2Icon /> Delete the view
                        </DropdownMenuItem>
                      </>
                    )}
                  </DropdownMenuContent>
                </DropdownMenu>
              )}
            </div>
          );
        })}
        <DropdownMenu>
          <DropdownMenuTrigger className="icon-btn mb-0.5 size-8" aria-label="New view">
            <PlusIcon />
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start">
            {(Object.keys(LAYOUT_LABELS) as Layout[]).map((l) => {
              const Icon = LAYOUT_ICONS[l];
              return (
                <DropdownMenuItem key={l} onSelect={() => addView(l)}>
                  <Icon /> {LAYOUT_LABELS[l]}
                </DropdownMenuItem>
              );
            })}
          </DropdownMenuContent>
        </DropdownMenu>
      </div>
      <Toolbar entry={entry} view={view} onChange={setView} count={shown.length} total={data.rows.length} onNew={() => void newRow()} />
      <div className="min-h-0 flex-1 overflow-auto">
        {view.layout === "board" ? (
          <Board entry={entry} view={view} rows={shown} onOpen={setOpen} onMove={save} onNew={(values) => void newRow(values)} />
        ) : view.layout === "list" ? (
          <List entry={entry} rows={shown} onOpen={setOpen} />
        ) : (
          <Grid columns={columns} view={view} rows={shown} onSave={save} onOpen={setOpen} onSort={(sort) => setView({ sort })} onNew={() => void newRow()} />
        )}
      </div>
      <RowPeek entry={entry} table={table} index={open} onClose={() => setOpen(undefined)} />
    </div>
  );
}

/* ----------------------------------- Toolbar ----------------------------------- */

function Toolbar({ entry, view, onChange, count, total, onNew }: { entry: TableEntry; view: View; onChange: (v: Partial<View>) => void; count: number; total: number; onNew: () => void }) {
  const setFilter = (i: number, f?: Filter) => onChange({ filters: f ? view.filters.map((x, j) => (j === i ? f : x)) : view.filters.filter((_, j) => j !== i) });
  const newFilter = (c: Column): Filter => (c.type === "option" ? { column: c.name, values: [] } : c.type === "date" ? { column: c.name, date: "past" } : { column: c.name, text: "" });
  return (
    <div className="flex flex-wrap items-center gap-1.5 px-4 py-2.5">
      {view.filters.map((f, i) => {
        const c = entry.columns.find((x) => x.name === f.column);
        return c ? <FilterChip key={i} column={c} filter={f} onChange={(next) => setFilter(i, next)} /> : null;
      })}
      <DropdownMenu>
        <DropdownMenuTrigger className="chip h-7 border-dashed text-muted-foreground">
          <PlusIcon /> Filter
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="w-52">
          {entry.columns.map((c) => {
            const Icon = TYPE_ICONS[c.type];
            return (
              <DropdownMenuItem key={c.name} onSelect={() => onChange({ filters: [...view.filters, newFilter(c)] })}>
                <Icon /> {c.name}
              </DropdownMenuItem>
            );
          })}
        </DropdownMenuContent>
      </DropdownMenu>
      <span className="ml-auto flex items-center gap-1.5">
        <span className="text-[12px] text-muted-foreground tabular-nums">{count === total ? `${total} rows` : `${count} of ${total}`}</span>
        <label className="flex h-7 items-center gap-1.5 rounded-md border bg-card px-2 text-[12.5px]">
          <SearchIcon className="size-3.5 text-muted-foreground" />
          <input value={view.search ?? ""} onChange={(e) => onChange({ search: e.target.value })} placeholder="Search" className="w-28 bg-transparent outline-none" />
        </label>
        {view.layout === "board" && (
          <DropdownMenu>
            <DropdownMenuTrigger className="chip h-7">
              <KanbanIcon /> {view.group ?? "Group by"}
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end">
              {entry.columns.filter((c) => c.type === "option").map((c) => (
                <DropdownMenuCheckboxItem key={c.name} checked={view.group === c.name} onCheckedChange={() => onChange({ group: c.name })}>
                  {c.name}
                </DropdownMenuCheckboxItem>
              ))}
            </DropdownMenuContent>
          </DropdownMenu>
        )}
        <DropdownMenu>
          <DropdownMenuTrigger className="chip h-7" aria-label="Shown columns">
            <EyeOffIcon /> {view.hidden?.length ? `${view.hidden.length} hidden` : "Columns"}
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="w-52">
            {entry.columns.map((c) => (
              <DropdownMenuCheckboxItem key={c.name} checked={!view.hidden?.includes(c.name)} onCheckedChange={(on) => onChange({ hidden: on ? view.hidden?.filter((h) => h !== c.name) : [...(view.hidden ?? []), c.name] })} onSelect={(e) => e.preventDefault()}>
                {c.name}
              </DropdownMenuCheckboxItem>
            ))}
          </DropdownMenuContent>
        </DropdownMenu>
        <Button size="sm" className="h-7" onClick={onNew}>
          <PlusIcon /> New
        </Button>
      </span>
    </div>
  );
}

/** A filter as a chip: an option column's values to tick, a text to contain, a date period. */
function FilterChip({ column, filter, onChange }: { column: Column; filter: Filter; onChange: (f?: Filter) => void }) {
  return (
    <span className="chip h-7 gap-1 pr-1">
      <DropdownMenu>
        <DropdownMenuTrigger className="flex items-center gap-1 outline-none">
          <span className="font-medium">{filterText(filter).replace(/ is $/, " is any")}</span>
          <ChevronDownIcon className="size-3" />
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="w-56">
          {column.type === "option" &&
            column.options?.map((o) => (
              <DropdownMenuCheckboxItem key={o.value} checked={filter.values?.includes(o.value)} onSelect={(e) => e.preventDefault()} onCheckedChange={(on) => onChange({ ...filter, values: on ? [...(filter.values ?? []), o.value] : (filter.values ?? []).filter((v) => v !== o.value) })}>
                <Badge small color={o.color}>{o.value}</Badge>
              </DropdownMenuCheckboxItem>
            ))}
          {column.type === "date" &&
            (Object.keys(DATE_RANGES) as DateRange[]).map((r) => (
              <DropdownMenuCheckboxItem key={r} checked={filter.date === r} onCheckedChange={() => onChange({ ...filter, date: r })}>
                {DATE_RANGES[r]}
              </DropdownMenuCheckboxItem>
            ))}
          {(column.type === "text" || column.type === "longtext") && (
            <div className="p-1.5">
              <input autoFocus value={filter.text ?? ""} onChange={(e) => onChange({ ...filter, text: e.target.value })} onKeyDown={(e) => e.stopPropagation()} placeholder="Contains…" className="h-8 w-full rounded-md border bg-card px-2 text-[13px] outline-none" />
            </div>
          )}
        </DropdownMenuContent>
      </DropdownMenu>
      <button type="button" className="rounded p-0.5 hover:bg-muted" aria-label="Remove the filter" onClick={() => onChange(undefined)}>
        <XIcon className="size-3" />
      </button>
    </span>
  );
}

/* ------------------------------------ Cells ------------------------------------ */

const optionColor = (c: Column, v: string) => c.options?.find((o) => o.value === v)?.color;

/** A value shown (an option as its badge, a date short, a long text on two lines). */
function Value({ column, value }: { column: Column; value: string }) {
  if (!value) return <span className="text-muted-foreground/50">—</span>;
  if (column.type === "option") return <Badge color={optionColor(column, value) ?? "gray"}>{value}</Badge>;
  if (column.type === "date") return <span className="tabular-nums">{new Date(`${value}T00:00:00`).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" })}</span>;
  if (column.type === "longtext") return <span className="line-clamp-2 text-[12.5px] whitespace-pre-line text-muted-foreground">{value}</span>;
  return <span className="truncate">{value}</span>;
}

/** A value edited in place: committed on blur / Enter (an option or a date as soon as it is chosen). */
function Input({ column, value, onSave, form }: { column: Column; value: string; onSave: (v: string) => void; form?: boolean }) {
  const [text, setText] = useState(value);
  useEffect(() => setText(value), [value]);
  const commit = () => text !== value && onSave(text);
  if (column.type === "option")
    return (
      <DropdownMenu>
        <DropdownMenuTrigger className="flex min-h-7 w-full items-center outline-none">
          <Value column={column} value={value} />
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="w-52">
          {column.options?.map((o) => (
            <DropdownMenuItem key={o.value} onSelect={() => onSave(o.value)}>
              <Badge small color={o.color}>{o.value}</Badge>
              {o.value === value && <CheckIcon className="ml-auto" />}
            </DropdownMenuItem>
          ))}
          {value && (
            <>
              <DropdownMenuSeparator />
              <DropdownMenuItem onSelect={() => onSave("")}>Clear</DropdownMenuItem>
            </>
          )}
        </DropdownMenuContent>
      </DropdownMenu>
    );
  if (column.type === "date") return <input type="date" value={value} onChange={(e) => onSave(e.target.value)} className={cn("h-7 w-full bg-transparent text-[13px] outline-none", form && "h-9 rounded-md border bg-card px-2", !value && !form && "text-transparent focus:text-foreground")} />;
  if (column.type === "longtext" && form) return <textarea value={text} onChange={(e) => setText(e.target.value)} onBlur={commit} rows={6} className="w-full rounded-md border bg-card p-2 text-[13px] outline-none focus:ring-2 focus:ring-ring/30" />;
  return <input value={text} onChange={(e) => setText(e.target.value)} onBlur={commit} onKeyDown={(e) => e.key === "Enter" && (e.currentTarget as HTMLInputElement).blur()} className={cn("h-7 w-full bg-transparent text-[13.5px] outline-none", form && "h-9 rounded-md border bg-card px-2")} />;
}

const Head = ({ column }: { column: Column }) => {
  const Icon = TYPE_ICONS[column.type];
  return (
    <span className="flex items-center gap-1.5 truncate">
      <Icon className="size-3.5 shrink-0 opacity-60" aria-hidden />
      {column.name}
    </span>
  );
};

/* ------------------------------------ Layouts ----------------------------------- */

const WIDTHS: Record<ColumnType, number> = { text: 220, longtext: 300, date: 130, option: 140 };

function Grid({ columns, view, rows, onSave, onOpen, onSort, onNew }: { columns: Column[]; view: View; rows: { row: Row; index: number }[]; onSave: (index: number, values: Row) => void; onOpen: (index: number) => void; onSort: (s: View["sort"]) => void; onNew: () => void }) {
  return (
    <div className="px-2 pb-4">
      <table className="w-max min-w-full border-collapse text-[13.5px]">
        <thead className="sticky top-0 z-10 bg-background">
          <tr className="border-b">
            <th className="w-8" />
            {columns.map((c) => {
              const sorted = view.sort?.column === c.name;
              return (
                <th key={c.name} style={{ minWidth: WIDTHS[c.type], maxWidth: WIDTHS[c.type] * 1.5 }} className="h-9 border-r px-2 text-left text-[12.5px] font-medium text-muted-foreground last:border-r-0">
                  <button type="button" className="flex w-full items-center gap-1 hover:text-foreground" onClick={() => onSort(sorted && view.sort?.desc ? undefined : { column: c.name, desc: sorted })}>
                    <Head column={c} />
                    {sorted && (view.sort?.desc ? <ArrowDownIcon className="size-3" /> : <ArrowUpIcon className="size-3" />)}
                  </button>
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {rows.map(({ row, index }) => (
            <tr key={index} className="group border-b hover:bg-muted/40">
              <td className="w-8 text-center">
                <button type="button" className="icon-btn size-6 opacity-0 group-hover:opacity-100" aria-label="Open the row" onClick={() => onOpen(index)}>
                  <Maximize2Icon />
                </button>
              </td>
              {columns.map((c) => (
                <td key={c.name} style={{ maxWidth: WIDTHS[c.type] * 1.5 }} className="h-10 border-r px-2 align-middle last:border-r-0" onDoubleClick={() => c.type === "longtext" && onOpen(index)}>
                  {c.type === "longtext" ? <Value column={c} value={row[c.name] ?? ""} /> : <Input column={c} value={row[c.name] ?? ""} onSave={(v) => onSave(index, { [c.name]: v })} />}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      <button type="button" onClick={onNew} className="flex h-9 w-full items-center gap-1.5 px-3 text-[13px] text-muted-foreground hover:bg-muted/40">
        <PlusIcon className="size-3.5" /> New
      </button>
    </div>
  );
}

/** A card of the board or the list: the first text column as its title, the other values under it. */
const titleColumn = (entry: TableEntry) => entry.columns.find((c) => c.type === "text") ?? entry.columns[0];

function Board({ entry, view, rows, onOpen, onMove, onNew }: { entry: TableEntry; view: View; rows: { row: Row; index: number }[]; onOpen: (i: number) => void; onMove: (index: number, values: Row) => void; onNew: (values: Row) => void }) {
  const group = entry.columns.find((c) => c.name === view.group && c.type === "option") ?? entry.columns.find((c) => c.type === "option");
  const [dragged, setDragged] = useState<number>();
  if (!group) return <Empty icon={KanbanIcon}>A board groups the rows by an option column: add one in the Columns tab.</Empty>;
  const title = titleColumn(entry);
  const others = entry.columns.filter((c) => c !== title && c !== group && !view.hidden?.includes(c.name) && c.type !== "longtext").slice(0, 3);
  const lanes = [...(group.options ?? []).map((o) => ({ value: o.value, color: o.color as Color | undefined })), { value: "", color: undefined }];
  return (
    <div className="flex h-full gap-3 px-4 pb-4">
      {lanes.map((lane) => {
        const cards = rows.filter(({ row }) => (row[group.name] ?? "") === lane.value);
        if (!lane.value && !cards.length) return null;
        return (
          <div key={lane.value || "none"} className="flex w-68 shrink-0 flex-col gap-2 rounded-xl bg-muted/50 p-2" onDragOver={(e) => e.preventDefault()} onDrop={() => dragged !== undefined && (onMove(dragged, { [group.name]: lane.value }), setDragged(undefined))}>
            <div className="flex items-center gap-2 px-1 text-[12.5px]">
              {lane.value ? <Badge color={lane.color}>{lane.value}</Badge> : <span className="text-muted-foreground">No {group.name}</span>}
              <span className="text-muted-foreground">{cards.length}</span>
            </div>
            <div className="flex min-h-0 flex-col gap-2 overflow-y-auto">
              {cards.map(({ row, index }) => (
                <button key={index} type="button" draggable onDragStart={() => setDragged(index)} onClick={() => onOpen(index)} className="card flex flex-col gap-2 p-3 text-left text-[13px] hover:border-primary/40">
                  <span className="font-medium">{row[title.name] || "Untitled"}</span>
                  {others.some((c) => row[c.name]) && (
                    <span className="flex flex-wrap items-center gap-1.5 text-[12px] text-muted-foreground">
                      {others.filter((c) => row[c.name]).map((c) => (c.type === "option" ? <Badge key={c.name} small color={optionColor(c, row[c.name])}>{row[c.name]}</Badge> : <span key={c.name} className="flex items-center gap-1">{c.type === "date" && <CalendarIcon className="size-3" />}<Value column={c} value={row[c.name]} /></span>))}
                    </span>
                  )}
                </button>
              ))}
            </div>
            {lane.value && (
              <button type="button" onClick={() => onNew({ [group.name]: lane.value })} className="flex items-center gap-1 rounded-md px-2 py-1 text-[12.5px] text-muted-foreground hover:bg-card">
                <PlusIcon className="size-3.5" /> New
              </button>
            )}
          </div>
        );
      })}
    </div>
  );
}

function List({ entry, rows, onOpen }: { entry: TableEntry; rows: { row: Row; index: number }[]; onOpen: (i: number) => void }) {
  const title = titleColumn(entry);
  const option = entry.columns.find((c) => c.type === "option");
  const date = entry.columns.find((c) => c.type === "date");
  const note = entry.columns.find((c) => c.type === "longtext");
  return (
    <div className="mx-4 mb-4 flex flex-col divide-y rounded-xl border bg-card">
      {rows.map(({ row, index }) => (
        <button key={index} type="button" onClick={() => onOpen(index)} className="flex items-start gap-3 px-4 py-3 text-left hover:bg-muted/40">
          {option && <span className="w-28 shrink-0 pt-0.5"><Value column={option} value={row[option.name] ?? ""} /></span>}
          <span className="flex min-w-0 flex-1 flex-col gap-0.5">
            <span className="text-[13.5px] font-medium">{row[title.name] || "Untitled"}</span>
            {note && row[note.name] && <span className="line-clamp-1 text-[12.5px] text-muted-foreground">{row[note.name]}</span>}
          </span>
          {date && <span className="shrink-0 text-[12.5px] text-muted-foreground"><Value column={date} value={row[date.name] ?? ""} /></span>}
        </button>
      ))}
      {!rows.length && <p className="px-4 py-6 text-center text-[13px] text-muted-foreground">No row in this view.</p>}
    </div>
  );
}

/* ------------------------------------ The row ------------------------------------ */

/** A row opened on the side: every column as a field of a form, saved as each one is left. */
function RowPeek({ entry, table, index, onClose }: { entry: TableEntry; table: ExcelTable; index?: number; onClose: () => void }) {
  const row = index !== undefined ? table.data?.rows[index] : undefined;
  const title = titleColumn(entry);
  return (
    <Sheet open={!!row} onOpenChange={(o) => !o && onClose()}>
      <SheetContent className="w-[520px] sm:max-w-[520px]">
        {row && index !== undefined && (
          <>
            <SheetHeader>
              <SheetTitle className="text-[19px]">{row[title.name] || "Untitled"}</SheetTitle>
              <SheetDescription>Row {index + 1} of {entry.excel.name} · saved to Excel as you leave each field</SheetDescription>
            </SheetHeader>
            <div className="flex flex-col gap-4 overflow-y-auto px-4 pb-4">
              {entry.columns.map((c) => (
                <label key={c.name} className="grid grid-cols-[130px_1fr] items-start gap-3 text-[13.5px]">
                  <span className="pt-2 text-muted-foreground">
                    <Head column={c} />
                  </span>
                  {c.type === "option" ? (
                    <span className="flex flex-wrap gap-1 pt-1">
                      {c.options?.map((o) => (
                        <button key={o.value} type="button" onClick={() => void withToast("Saving…", table.update([{ index, values: { [c.name]: row[c.name] === o.value ? "" : o.value } }]))} className={cn("transition-opacity", row[c.name] === o.value ? "" : "opacity-35 hover:opacity-70")}>
                          <Badge color={o.color}>{o.value}</Badge>
                        </button>
                      ))}
                    </span>
                  ) : (
                    <Input form column={c} value={row[c.name] ?? ""} onSave={(v) => void withToast("Saving…", table.update([{ index, values: { [c.name]: v } }]))} />
                  )}
                </label>
              ))}
              <Button variant="ghost" className="self-start text-destructive" onClick={() => confirm("Delete this row from the Excel file?") && void withToast("Deleting…", table.remove([index]).then(onClose), "Row deleted.")}>
                <Trash2Icon /> Delete the row
              </Button>
            </div>
          </>
        )}
      </SheetContent>
    </Sheet>
  );
}

/* ------------------------------------ Columns ------------------------------------ */

/** The columns: order, name, type among the 4, options and their colors; each change written to the file (lists, colors, formatting). */
function ColumnsPanel({ entry, table }: { entry: TableEntry; table: ExcelTable }) {
  const [columns, setColumns] = useState(entry.columns);
  useEffect(() => setColumns(entry.columns), [entry.columns]);
  const [dragged, setDragged] = useState<number>();
  const commit = (next: Column[], removed?: string[]) => (setColumns(next), void withToast("Writing the columns to Excel…", table.setColumns(next, removed), "Columns written to Excel."));
  const change = (c: Column, values: Partial<Column>) => commit(columns.map((x) => (x === c ? { ...x, ...values } : x)));
  const values = (c: Column) => [...new Set((table.data?.rows ?? []).map((r) => r[c.name]).filter(Boolean))];
  return (
    <div className="min-h-0 flex-1 overflow-y-auto">
      <div className="mx-auto flex max-w-3xl flex-col gap-5 px-6 py-6">
        <Section
          title="Columns"
          description="Their order, names and types are written to the Excel file: an option column gets a drop-down list and a color per option there too. Drag to reorder."
          actions={
            <>
              <Button size="sm" variant="outline" onClick={() => void withToast("Formatting the table…", table.sync(), "Formatting synced.")}>
                <PaintbrushIcon /> Sync the formatting
              </Button>
              <Button size="sm" onClick={() => { const name = prompt("Name of the new column")?.trim(); if (name && !columns.some((c) => c.name === name)) commit([...columns, { name, type: "text" }]); }}>
                <PlusIcon /> Column
              </Button>
            </>
          }
        >
          <div className="flex flex-col divide-y">
            {columns.map((c, i) => {
              const Icon = TYPE_ICONS[c.type];
              return (
                <div key={c.was ?? c.name} className={cn("flex flex-col gap-2 py-3", dragged === i && "opacity-50")} draggable onDragStart={() => setDragged(i)} onDragOver={(e) => e.preventDefault()} onDrop={() => { if (dragged === undefined || dragged === i) return; const next = [...columns]; next.splice(i, 0, ...next.splice(dragged, 1)); setDragged(undefined); commit(next); }}>
                  <div className="flex items-center gap-2 text-[13.5px]">
                    <GripVerticalIcon className="size-4 cursor-grab text-muted-foreground" />
                    <Icon className="size-4 text-muted-foreground" />
                    <input defaultValue={c.name} key={c.name} onBlur={(e) => { const name = e.target.value.trim(); if (name && name !== c.name && !columns.some((x) => x.name === name)) change(c, { name, was: c.was ?? c.name }); else e.target.value = c.name; }} className="h-8 flex-1 rounded-md border border-transparent bg-transparent px-2 font-medium outline-none hover:border-input focus:border-input" />
                    <DropdownMenu>
                      <DropdownMenuTrigger className="chip h-8 w-32 justify-between">
                        {TYPE_LABELS[c.type]} <ChevronDownIcon />
                      </DropdownMenuTrigger>
                      <DropdownMenuContent align="end">
                        {COLUMN_TYPES.map((t) => {
                          const TIcon = TYPE_ICONS[t];
                          return (
                            <DropdownMenuItem key={t} onSelect={() => change(c, { type: t, options: t === "option" ? (c.options?.length ? c.options : values(c).slice(0, 12).map((value, k) => ({ value, color: COLORS[(k + 1) % COLORS.length] }))) : c.options })}>
                              <TIcon /> {TYPE_LABELS[t]}
                            </DropdownMenuItem>
                          );
                        })}
                      </DropdownMenuContent>
                    </DropdownMenu>
                    <button type="button" className="icon-btn" aria-label={`Delete ${c.name}`} onClick={() => confirm(`Delete the column "${c.name}" and its values from the Excel file?`) && commit(columns.filter((x) => x !== c), [c.was ?? c.name])}>
                      <Trash2Icon />
                    </button>
                  </div>
                  {c.type === "option" && <Options column={c} onChange={(options) => change(c, { options })} />}
                </div>
              );
            })}
          </div>
        </Section>
      </div>
    </div>
  );
}

/** The options of a column: each a badge whose color is picked in its menu; a new one typed in. */
function Options({ column, onChange }: { column: Column; onChange: (o: NonNullable<Column["options"]>) => void }) {
  const options = column.options ?? [];
  const [text, setText] = useState("");
  return (
    <div className="flex flex-wrap items-center gap-1.5 pl-12">
      {options.map((o) => (
        <DropdownMenu key={o.value}>
          <DropdownMenuTrigger className="outline-none">
            <Badge color={o.color}>{o.value}</Badge>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start" className="w-44">
            <div className="grid grid-cols-5 gap-1.5 p-2">
              {COLORS.map((color) => (
                <button key={color} type="button" aria-label={color} onClick={() => onChange(options.map((x) => (x === o ? { ...x, color } : x)))} className={cn("size-6 rounded-full", SWATCHES[color], o.color === color && "ring-2 ring-ring ring-offset-2")} />
              ))}
            </div>
            <DropdownMenuSeparator />
            <DropdownMenuItem onSelect={() => { const value = prompt("Option", o.value)?.trim(); if (value && !options.some((x) => x.value === value)) onChange(options.map((x) => (x === o ? { ...x, value } : x))); }}>
              <PencilIcon /> Rename
            </DropdownMenuItem>
            <DropdownMenuItem className="text-destructive" onSelect={() => onChange(options.filter((x) => x !== o))}>
              <Trash2Icon /> Remove
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      ))}
      <input
        value={text}
        onChange={(e) => setText(e.target.value.replace(/,/g, ""))}
        onKeyDown={(e) => {
          const value = text.trim();
          if (e.key === "Enter" && value && !options.some((o) => o.value === value)) (onChange([...options, { value, color: COLORS[(options.length + 1) % COLORS.length] }]), setText(""));
        }}
        placeholder="+ option"
        className="h-6 w-28 rounded-full border border-dashed bg-transparent px-2.5 text-[12px] outline-none"
      />
    </div>
  );
}

