import { useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import {
  ArrowDownIcon,
  ArrowUpIcon,
  ExternalLinkIcon,
  CheckIcon,
  CloudAlertIcon,
  FileSpreadsheetIcon,
  FilterIcon,
  LoaderIcon,
  PanelLeftIcon,
  CloudIcon,
  FolderIcon,
  PlusIcon,
  SearchIcon,
  TableIcon,
  Trash2Icon,
} from "lucide-react";
import { z } from "zod";
import { errorResult, textResult, useAgentFeatures, useLatest } from "@/agent/server";
import { FieldInput, OptionBadges } from "@/components/field-input";
import { IconButton } from "@/components/icon-button";
import { EmptyState } from "@/components/page";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import { TINTS } from "@/lib/colors";
import type { ExcelFile } from "@/lib/excel";
import { labelOf } from "@/lib/schema";
import type { Row, Schema } from "@/lib/schema";
import { handles, links, schemas } from "@/lib/store";
import type { FileLink } from "@/lib/store";
import { useExcel } from "@/lib/use-excel";
import type { ImageStore } from "@/storage/types";
import { ImagesPicker } from "@/components/images-picker";
import type { ImagesChoice } from "@/components/images-picker";
import { MicrosoftSignIn } from "@/components/storage-ui";
import { cn } from "@/lib/utils";
import { applyView, describeView, useView } from "@/lib/view";
import type { View } from "@/lib/view";

export const Route = createFileRoute("/files/$id")({ component: FilePage });

function FilePage() {
  const { id } = Route.useParams();
  const link = links.use().find((l) => l.id === id);
  const schema = schemas.use().find((s) => s.id === link?.schemaId);
  const { state, reopen } = useExcel(link, schema);

  if (state.status === "loading")
    return <p className="shimmer p-8 text-[12px]">Opening the file…</p>;
  if (state.status === "permission")
    return (
      <div className="p-8">
        <EmptyState icon={<FileSpreadsheetIcon />} title={`Allow access to ${state.fileName}`}>
          <p className="mb-3">The browser needs your permission to read and write this file.</p>
          <Button size="sm" onClick={() => void reopen()}>
            Allow access
          </Button>
        </EmptyState>
      </div>
    );
  if (state.status === "signin")
    return (
      <div className="mx-auto max-w-md p-8">
        <EmptyState icon={<CloudIcon />} title="Sign in to open this file">
          <p className="mb-3">It is stored on SharePoint / OneDrive.</p>
          <MicrosoftSignIn onSignedIn={() => void reopen()} />
        </EmptyState>
      </div>
    );
  if (state.status === "error" || !link || !schema)
    return <p className="p-8 text-[12px] text-destructive">{state.status === "error" ? state.error : "Not found."}</p>;
  return <Editor link={link} schema={schema} excel={state.excel} images={state.images} onReload={() => void reopen()} />;
}

function Editor({
  link,
  schema,
  excel,
  images,
  onReload,
}: {
  link: FileLink;
  schema: Schema;
  excel: ExcelFile;
  images?: ImageStore;
  onReload: () => void;
}) {
  const [view, setView] = useView(link.id);
  const rows = applyView(excel.rows, view, schema);
  const [selected, setSelected] = useState<string | null>(null);
  const [error, setError] = useState("");
  const latest = useLatest({ rows, view });

  const run = (action: () => void) => {
    try {
      action();
      setError("");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };
  const change = (row: Row, field: string, value: unknown) => run(() => excel.update([{ id: row.id, values: { [field]: value } }]));
  const [adding, setAdding] = useState(false);
  const [imagesOpen, setImagesOpen] = useState(false);

  useAgent(link, schema, excel, latest);

  return (
    <div className="flex h-full flex-col">
      <Toolbar
        link={link}
        schema={schema}
        excel={excel}
        view={view}
        setView={setView}
        count={rows.length}
        onAdd={() => setAdding(true)}
        onImagesFolder={() => setImagesOpen(true)}
      />
      {(error || excel.save.status === "error") && (
        <p className="border-b bg-destructive/5 px-4 py-1.5 text-[12px] text-destructive">{error || excel.save.error}</p>
      )}
      {view.layout === "table" ? (
        <Table schema={schema} rows={rows} view={view} setView={setView} images={images} onChange={change} onOpen={(row) => {
          setSelected(row.id);
          setView({ layout: "form" });
        }} onDelete={(row) => run(() => excel.remove([row.id]))} />
      ) : (
        <FormLayout schema={schema} rows={rows} selected={selected} onSelect={setSelected} images={images} onChange={change} />
      )}
      <ImagesDialog link={link} open={imagesOpen} onClose={() => setImagesOpen(false)} onChanged={onReload} />
      <NewRowDialog
        open={adding}
        schema={schema}
        images={images}
        onClose={() => setAdding(false)}
        onCreate={(values) => {
          const [row] = excel.insert([values]);
          setSelected(row.id);
        }}
      />
    </div>
  );
}

/** What the agent can see and do on this page. */
function useAgent(link: FileLink, schema: Schema, excel: ExcelFile, latest: { current: { rows: Row[]; view: View } }) {
  const jsonl = (rows: Row[]) => rows.map((r) => JSON.stringify(r)).join("\n");
  const page = `File editor: "${link.name}" (${link.source.kind === "sharepoint" ? "SharePoint / OneDrive" : "local"} file ${link.source.name}, schema "${schema.name}"${schema.description ? `: ${schema.description}` : ""}).
${excel.rows.length} rows in total; the user's current view shows ${latest.current.rows.length} rows (${describeView(latest.current.view)}).
excel://view holds the rows of the current view, excel://rows all rows (one JSON row per line, with its id).
You may add, update and delete rows.`;

  useAgentFeatures(
    page,
    (server) => {
      const items = [
        server.registerResource(
          "schema",
          "excel://schema",
          { title: "Fields of the rows", mimeType: "application/json", annotations: { priority: 1 } },
          async (uri) => ({ contents: [{ uri: uri.href, text: JSON.stringify(schema.fields) }] }),
        ),
        server.registerResource(
          "view",
          "excel://view",
          { title: "Rows of the current view", mimeType: "application/jsonl", annotations: { priority: 1 } },
          async (uri) => ({ contents: [{ uri: uri.href, text: jsonl(latest.current.rows) }] }),
        ),
        server.registerResource(
          "rows",
          "excel://rows",
          { title: "All rows", mimeType: "application/jsonl", annotations: { priority: 0.3 } },
          async (uri) => ({ contents: [{ uri: uri.href, text: jsonl(excel.rows) }] }),
        ),
        server.registerPrompt("summary", { title: "Summarize the rows of this view" }, () => ({
          messages: [{ role: "user", content: { type: "text", text: "Summarize the rows of the current view: counts, notable values, missing data." } }],
        })),
      ];
      const values = z.record(z.string(), z.unknown()).describe("field name → value (options: allowed values; dates: YYYY-MM-DD)");
      return [
        ...items,
        server.registerTool(
          "add_rows",
          { description: "Adds rows at the end of the file.", inputSchema: z.object({ rows: z.array(values) }) },
          async ({ rows }) => {
            try {
              const added = excel.insert(rows);
              return textResult({ addedIds: added.map((r) => r.id) }, { name: "delete_rows", arguments: { ids: added.map((r) => r.id) } });
            } catch (error) {
              return errorResult(error);
            }
          },
        ),
        server.registerTool(
          "update_rows",
          {
            description: "Changes some values of existing rows (only the given fields change).",
            inputSchema: z.object({ updates: z.array(z.object({ id: z.string(), values })) }),
          },
          async ({ updates }) => {
            try {
              const before = excel.update(updates);
              return textResult({ updatedIds: updates.map((u) => u.id) }, { name: "update_rows", arguments: { updates: before } });
            } catch (error) {
              return errorResult(error);
            }
          },
        ),
        server.registerTool(
          "delete_rows",
          { description: "Deletes rows by id.", inputSchema: z.object({ ids: z.array(z.string()) }) },
          async ({ ids }) => {
            try {
              const removed = excel.remove(ids);
              return textResult({ deletedIds: ids }, { name: "add_rows", arguments: { rows: removed } });
            } catch (error) {
              return errorResult(error);
            }
          },
        ),
      ];
    },
    [link.id, schema, excel],
  );
}

/** SharePoint / OneDrive: the file in Excel online. Local: the file itself, which the system opens in Excel. */
async function openFile(link: FileLink) {
  if (link.source.kind === "sharepoint") return void window.open(link.source.webUrl, "_blank", "noopener");
  const handle = await handles.file(link.id);
  if (!handle) return;
  const url = URL.createObjectURL(await handle.getFile());
  Object.assign(document.createElement("a"), { href: url, download: link.source.name }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

function Toolbar({
  link,
  schema,
  excel,
  view,
  setView,
  count,
  onAdd,
  onImagesFolder,
}: {
  link: FileLink;
  schema: Schema;
  excel: ExcelFile;
  view: View;
  setView: (patch: Partial<View>) => void;
  count: number;
  onAdd: () => void;
  onImagesFolder: () => void;
}) {
  const optionFields = schema.fields.filter((f) => f.type === "option");
  const active = Object.values(view.filters).filter((v) => v.length).length;
  return (
    <div className="flex h-11 shrink-0 items-center gap-1.5 border-b px-3">
      <span className="mr-1 truncate font-medium">{link.name}</span>
      <span className="text-[12px] text-muted-foreground tabular-nums">
        {count === excel.rows.length ? count : `${count} / ${excel.rows.length}`} rows
      </span>
      <SaveState excel={excel} />
      <span className="ml-auto" />
      <label className="flex h-7 items-center gap-1.5 rounded-lg border bg-background px-2 transition-shadow focus-within:border-ring focus-within:ring-3 focus-within:ring-ring/20">
        <SearchIcon className="size-3.5 text-muted-foreground" />
        <input value={view.search} onChange={(e) => setView({ search: e.target.value })} placeholder="Search" className="w-32 bg-transparent outline-none" />
      </label>
      {optionFields.length > 0 && (
        <Popover>
          <PopoverTrigger render={<Button variant="ghost" size="sm" className={cn("text-muted-foreground", active && "text-foreground")} />}>
            <FilterIcon /> {active ? `${active} filter${active > 1 ? "s" : ""}` : "Filter"}
          </PopoverTrigger>
          <PopoverContent align="end" className="max-h-96 w-60 gap-3 overflow-y-auto">
            {optionFields.map((field) => (
              <div key={field.name} className="flex flex-col gap-1">
                <span className="text-[11px] font-medium text-muted-foreground uppercase">{labelOf(field)}</span>
                {field.options?.map((option) => {
                  const selected = view.filters[field.name] ?? [];
                  const on = selected.includes(option.value);
                  return (
                    <label key={option.value} className="flex items-center gap-2 py-0.5">
                      <Checkbox
                        checked={on}
                        onCheckedChange={() =>
                          setView({ filters: { ...view.filters, [field.name]: on ? selected.filter((v) => v !== option.value) : [...selected, option.value] } })
                        }
                      />
                      <span className={cn("inline-flex h-5 items-center rounded-full px-2 text-[12px]", TINTS[option.color])}>{option.value}</span>
                    </label>
                  );
                })}
              </div>
            ))}
            {active > 0 && (
              <Button variant="ghost" size="xs" className="w-fit" onClick={() => setView({ filters: {} })}>
                Clear filters
              </Button>
            )}
          </PopoverContent>
        </Popover>
      )}
      <Segmented
        value={view.layout}
        onChange={(layout) => setView({ layout })}
        items={[
          { value: "table", label: "Table", icon: <TableIcon /> },
          { value: "form", label: "Form", icon: <PanelLeftIcon /> },
        ]}
      />
      <IconButton label={link.images ? `Images folder: ${link.images.name} (change)` : "Choose an images folder"} onClick={onImagesFolder}>
        <FolderIcon />
      </IconButton>
      <IconButton label="Add row" onClick={onAdd}>
        <PlusIcon />
      </IconButton>
      <IconButton label="Open the Excel file" onClick={() => void openFile(link)}>
        <ExternalLinkIcon />
      </IconButton>
    </div>
  );
}

function Segmented<T extends string>({ value, onChange, items }: { value: T; onChange: (v: T) => void; items: { value: T; label: string; icon: React.ReactNode }[] }) {
  return (
    <div className="flex rounded-lg bg-muted p-0.5">
      {items.map((item) => (
        <Tooltip key={item.value}>
          <TooltipTrigger
            onClick={() => onChange(item.value)}
            aria-label={item.label}
            className={cn(
              "flex size-6 items-center justify-center rounded-md text-muted-foreground transition-all [&_svg]:size-3.5",
              value === item.value && "bg-background text-foreground shadow-soft",
            )}
          >
            {item.icon}
          </TooltipTrigger>
          <TooltipContent>{item.label}</TooltipContent>
        </Tooltip>
      ))}
    </div>
  );
}

function SaveState({ excel }: { excel: ExcelFile }) {
  const { status } = excel.save;
  return (
    <span className="ml-1 flex items-center gap-1 text-[11px] text-muted-foreground">
      {status === "saving" ? (
        <LoaderIcon className="size-3 animate-spin" />
      ) : status === "error" ? (
        <CloudAlertIcon className="size-3 text-destructive" />
      ) : (
        <CheckIcon className="size-3" />
      )}
      {status === "saving" ? "Saving" : status === "error" ? "Not saved" : "Saved"}
    </span>
  );
}

function Table({
  schema,
  rows,
  view,
  setView,
  images,
  onChange,
  onOpen,
  onDelete,
}: {
  schema: Schema;
  rows: Row[];
  view: View;
  setView: (patch: Partial<View>) => void;
  images?: ImageStore;
  onChange: (row: Row, field: string, value: unknown) => void;
  onOpen: (row: Row) => void;
  onDelete: (row: Row) => void;
}) {
  const sortBy = (field: string) =>
    setView({
      sort: view.sort?.field !== field ? { field, direction: "asc" } : view.sort.direction === "asc" ? { field, direction: "desc" } : null,
    });
  if (rows.length === 0)
    return <p className="p-8 text-center text-[12px] text-muted-foreground">No rows{view.search || Object.keys(view.filters).length ? " match this view" : " yet"}.</p>;
  return (
    <div className="min-h-0 flex-1 overflow-auto">
      <table className="w-max min-w-full border-separate border-spacing-0">
        <thead className="sticky top-0 z-10 bg-background/95 backdrop-blur">
          <tr>
            {schema.fields.map((field) => (
              <th key={field.name} className="border-b border-r px-0 text-left font-medium last:border-r-0">
                <button
                  type="button"
                  onClick={() => sortBy(field.name)}
                  title={field.description}
                  className="flex h-8 w-full items-center gap-1 px-2 text-[12px] text-muted-foreground hover:text-foreground"
                >
                  {labelOf(field)}
                  {field.required && <span className="text-destructive/70">*</span>}
                  {view.sort?.field === field.name &&
                    (view.sort.direction === "asc" ? <ArrowUpIcon className="size-3" /> : <ArrowDownIcon className="size-3" />)}
                </button>
              </th>
            ))}
            <th className="w-16 border-b" />
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.id} className="group hover:bg-muted/40">
              {schema.fields.map((field) => (
                <td
                  key={field.name}
                  className={cn(
                    "border-b border-r p-0 align-top last:border-r-0",
                    field.type === "text" ? "min-w-72 max-w-96" : field.type === "string" ? "min-w-44" : "min-w-24",
                  )}
                >
                  <FieldInput variant="cell" field={field} value={row[field.name]} images={images} onChange={(v) => onChange(row, field.name, v)} />
                </td>
              ))}
              <td className="border-b px-1 align-middle">
                <div className="flex opacity-0 transition-opacity group-hover:opacity-100">
                  <IconButton label="Open in form" size="icon-xs" onClick={() => onOpen(row)}>
                    <PanelLeftIcon />
                  </IconButton>
                  <IconButton label="Delete row" size="icon-xs" onClick={() => onDelete(row)}>
                    <Trash2Icon />
                  </IconButton>
                </div>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function FormLayout({
  schema,
  rows,
  selected,
  onSelect,
  images,
  onChange,
}: {
  schema: Schema;
  rows: Row[];
  selected: string | null;
  onSelect: (id: string) => void;
  images?: ImageStore;
  onChange: (row: Row, field: string, value: unknown) => void;
}) {
  const titleField = schema.fields.find((f) => f.type === "string") ?? schema.fields[0];
  const row = rows.find((r) => r.id === selected) ?? rows[0];
  const optionFields = schema.fields.filter((f) => f.type === "option");
  return (
    <div className="flex min-h-0 flex-1">
      <ul className="flex w-80 shrink-0 flex-col gap-1.5 overflow-y-auto border-r bg-muted/30 p-2">
        {rows.map((r) => {
          const tags = optionFields.filter((f) => [r[f.name]].flat().some((v) => v !== null && v !== undefined && v !== ""));
          return (
            <li key={r.id}>
              <button
                type="button"
                onClick={() => onSelect(r.id)}
                className={cn(
                  "flex w-full flex-col items-start gap-1.5 rounded-lg border bg-card px-3 py-2 text-left shadow-soft transition-colors hover:border-ring/40",
                  r.id === row?.id && "border-ring ring-3 ring-ring/15",
                )}
              >
                <span className="w-full truncate font-medium">{String(r[titleField?.name] ?? "") || <span className="font-normal text-muted-foreground">Untitled</span>}</span>
                {tags.length > 0 && (
                  <span className="flex flex-wrap gap-1">
                    {tags.map((f) => (
                      <OptionBadges key={f.name} field={f} value={r[f.name]} />
                    ))}
                  </span>
                )}
              </button>
            </li>
          );
        })}
      </ul>
      <div className="min-w-0 flex-1 overflow-y-auto">
        {row ? (
          <div key={row.id} className="mx-auto flex max-w-xl flex-col gap-4 px-6 py-6 animate-in fade-in-0">
            {schema.fields.map((field) => (
              <label key={field.name} className="flex flex-col gap-1.5">
                <span className="text-[12px] font-medium">
                  {labelOf(field)}
                  {field.required && <span className="text-destructive/70"> *</span>}
                </span>
                <FieldInput variant="form" field={field} value={row[field.name]} images={images} onChange={(v) => onChange(row, field.name, v)} />
                {field.description && <span className="text-[11px] text-muted-foreground">{field.description}</span>}
              </label>
            ))}
            <p className="font-mono text-[11px] text-muted-foreground/60">{row.id}</p>
          </div>
        ) : (
          <p className="p-8 text-center text-[12px] text-muted-foreground">No row selected.</p>
        )}
      </div>
    </div>
  );
}

/** New row: a form with every field; required fields are checked on create. */
function NewRowDialog({
  open,
  schema,
  images,
  onClose,
  onCreate,
}: {
  open: boolean;
  schema: Schema;
  images?: ImageStore;
  onClose: () => void;
  onCreate: (values: Record<string, unknown>) => void;
}) {
  const [values, setValues] = useState<Record<string, unknown>>({});
  const [error, setError] = useState("");
  const close = () => {
    setValues({});
    setError("");
    onClose();
  };
  const create = () => {
    try {
      onCreate(values);
      close();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  };
  return (
    <Dialog open={open} onOpenChange={(o) => !o && close()}>
      <DialogContent className="max-h-[85dvh] grid-cols-[minmax(0,1fr)] gap-4 overflow-y-auto sm:max-w-lg">
        <DialogTitle className="text-[13px] font-medium">New row</DialogTitle>
        {schema.fields.map((field) => (
          <label key={field.name} className="flex flex-col gap-1.5">
            <span className="text-[12px] font-medium">
              {labelOf(field)}
              {field.required && <span className="text-destructive/70"> *</span>}
            </span>
            <FieldInput variant="form" field={field} value={values[field.name]} images={images} onChange={(v) => setValues((s) => ({ ...s, [field.name]: v }))} />
            {field.description && <span className="text-[11px] text-muted-foreground">{field.description}</span>}
          </label>
        ))}
        {error && <p className="text-[12px] text-destructive">{error}</p>}
        <div className="flex justify-end gap-2">
          <Button variant="ghost" onClick={close}>
            Cancel
          </Button>
          <Button onClick={create}>Create</Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}

/** Sets, changes or removes the images folder of a file (local or SharePoint). */
function ImagesDialog({ link, open, onClose, onChanged }: { link: FileLink; open: boolean; onClose: () => void; onChanged: () => void }) {
  const save = async (choice?: ImagesChoice) => {
    if (choice?.kind === "local") await handles.setImages(link.id, choice.handle);
    links.put({ ...link, images: choice?.kind === "local" ? { kind: "local", name: choice.handle.name } : choice });
    onClose();
    onChanged();
  };
  return (
    <Dialog open={open} onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="grid-cols-[minmax(0,1fr)] gap-3 sm:max-w-md">
        <DialogTitle className="text-[13px] font-medium">Images folder</DialogTitle>
        <p className="text-[12px] text-muted-foreground">
          {link.images ? `Current: ${link.images.name} (${link.images.kind === "sharepoint" ? "SharePoint, cells hold links" : "this computer, cells hold file names"}).` : "No images folder yet."}
        </p>
        <ImagesPicker onChange={(choice) => choice && void save(choice)} initialKind={link.source.kind} />
        {link.images && (
          <Button variant="ghost" size="sm" className="w-fit text-muted-foreground" onClick={() => void save(undefined)}>
            Remove the images folder
          </Button>
        )}
      </DialogContent>
    </Dialog>
  );
}
