import { useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { applyPatch } from "fast-json-patch";
import type { Operation } from "fast-json-patch";
import { ArrowDownIcon, ArrowUpIcon, AsteriskIcon, DownloadIcon, PlusIcon, Trash2Icon, XIcon } from "lucide-react";
import { z } from "zod";
import { errorResult, textResult, useAgentFeatures, useLatest } from "@/agent/server";
import { IconButton } from "@/components/icon-button";
import { Page } from "@/components/page";
import { SimpleSelect } from "@/components/simple-select";
import { Button } from "@/components/ui/button";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Input } from "@/components/ui/input";
import { Switch } from "@/components/ui/switch";
import { DOTS, TINTS } from "@/lib/colors";
import { COLORS, FIELD_TYPES, schemaSchema } from "@/lib/schema";
import type { Field, Schema } from "@/lib/schema";
import { links, schemas } from "@/lib/store";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/schemas/$id")({ component: SchemaPage });

const TYPE_LABELS: Record<Field["type"], string> = {
  string: "Short text",
  text: "Long text",
  number: "Number",
  integer: "Integer",
  boolean: "Yes / no",
  date: "Date",
  option: "Option",
  image: "Image",
};

function SchemaPage() {
  const { id } = Route.useParams();
  const schema = schemas.use().find((s) => s.id === id);
  const navigate = useNavigate();
  const latest = useLatest(schema);

  useAgentFeatures(
    schema
      ? `Schema editor for "${schema.name}". The schema (resource schema://current) defines the columns of Excel files: field name (column header), type, label, description, required, and for "option" fields the allowed values (with a color) and "multiple". Change it with patch_schema (JSON Patch on the schema object).`
      : "Schema not found.",
    (server) => [
      server.registerResource(
        "schema",
        "schema://current",
        { title: "Schema being edited", mimeType: "application/json", annotations: { priority: 1 } },
        async (uri) => ({ contents: [{ uri: uri.href, text: JSON.stringify(latest.current ?? null, null, 1) }] }),
      ),
      server.registerTool(
        "patch_schema",
        {
          description: `Changes the schema with JSON Patch operations (RFC 6902) applied in order, e.g. {"op":"add","path":"/fields/-","value":{"name":"price","type":"number"}} or {"op":"replace","path":"/fields/2/label","value":"Price"}. Field types: ${FIELD_TYPES.join(", ")}. Option colors: ${COLORS.join(", ")}. The id cannot change.`,
          inputSchema: z.object({
            operations: z.array(
              z.object({
                op: z.enum(["add", "remove", "replace", "move", "copy", "test"]),
                path: z.string(),
                value: z.unknown().optional(),
                from: z.string().optional(),
              }),
            ),
          }),
        },
        async ({ operations }) => {
          const before = latest.current;
          if (!before) return errorResult("No schema open.");
          try {
            const next = applyPatch(structuredClone(before), operations as Operation[], true, false).newDocument;
            const parsed = schemaSchema.parse({ ...next, id: before.id });
            schemas.put(parsed);
            latest.current = parsed;
            return textResult(`Schema updated (${parsed.fields.length} fields).`, {
              name: "patch_schema",
              arguments: { operations: [{ op: "replace", path: "", value: before }] },
            });
          } catch (error) {
            return errorResult(error instanceof z.ZodError ? z.prettifyError(error) : error);
          }
        },
      ),
      server.registerPrompt("suggest-fields", { title: "Suggest fields for this schema" }, () => ({
        messages: [{ role: "user", content: { type: "text", text: "Suggest and add useful fields for this schema, with good types and options." } }],
      })),
    ],
    [id],
  );

  if (!schema)
    return (
      <Page title="Schema not found">
        <span />
      </Page>
    );

  const save = (patch: Partial<Schema>) => schemas.put({ ...schema, ...patch });
  const setField = (index: number, patch: Partial<Field>) =>
    save({ fields: schema.fields.map((f, i) => (i === index ? { ...f, ...patch } : f)) });
  const move = (index: number, delta: number) => {
    const fields = [...schema.fields];
    const [field] = fields.splice(index, 1);
    fields.splice(index + delta, 0, field);
    save({ fields });
  };
  const check = schemaSchema.safeParse(schema);
  const used = links.all().filter((l) => l.schemaId === schema.id).length;

  return (
    <Page
      title={
        <input
          value={schema.name}
          onChange={(e) => save({ name: e.target.value })}
          className="w-full bg-transparent outline-none"
          aria-label="Schema name"
        />
      }
      subtitle={`${schema.fields.length} fields${used ? ` · used by ${used} file(s)` : ""}`}
      actions={
        <div className="flex gap-0.5">
          <IconButton
            label="Export JSON"
            onClick={() => {
              const url = URL.createObjectURL(new Blob([JSON.stringify(schema, null, 2)], { type: "application/json" }));
              Object.assign(document.createElement("a"), { href: url, download: `${schema.name}.json` }).click();
            }}
          >
            <DownloadIcon />
          </IconButton>
          <IconButton
            label="Delete schema"
            disabled={used > 0}
            onClick={() => {
              schemas.remove(schema.id);
              void navigate({ to: "/schemas" });
            }}
          >
            <Trash2Icon />
          </IconButton>
        </div>
      }
    >
      <div className="grid grid-cols-2 gap-3">
        <label className="flex flex-col gap-1">
          <span className="text-[12px] text-muted-foreground">Worksheet</span>
          <Input value={schema.sheet} maxLength={31} onChange={(e) => save({ sheet: e.target.value })} />
        </label>
        <label className="flex flex-col gap-1">
          <span className="text-[12px] text-muted-foreground">Description</span>
          <Input value={schema.description ?? ""} placeholder="What a row is (helps the agent)" onChange={(e) => save({ description: e.target.value })} />
        </label>
      </div>

      {!check.success && (
        <p className="rounded-md border border-destructive/20 bg-destructive/5 px-3 py-2 text-[12px] text-destructive whitespace-pre-wrap">
          {z.prettifyError(check.error)}
        </p>
      )}

      <div className="flex flex-col gap-2">
        {schema.fields.map((field, index) => (
          <FieldCard
            key={index}
            field={field}
            first={index === 0}
            last={index === schema.fields.length - 1}
            onChange={(patch) => setField(index, patch)}
            onMove={(delta) => move(index, delta)}
            onRemove={() => save({ fields: schema.fields.filter((_, i) => i !== index) })}
          />
        ))}
        <Button
          variant="ghost"
          size="sm"
          className="w-fit text-muted-foreground"
          onClick={() => save({ fields: [...schema.fields, { name: `field_${schema.fields.length + 1}`, type: "string" }] })}
        >
          <PlusIcon /> Add field
        </Button>
      </div>
    </Page>
  );
}

function FieldCard({
  field,
  first,
  last,
  onChange,
  onMove,
  onRemove,
}: {
  field: Field;
  first: boolean;
  last: boolean;
  onChange: (patch: Partial<Field>) => void;
  onMove: (delta: number) => void;
  onRemove: () => void;
}) {
  return (
    <div className="group flex flex-col gap-2 rounded-xl border bg-card p-2.5 shadow-soft transition-shadow hover:shadow-float">
      <div className="flex items-center gap-2">
        <Input
          value={field.name}
          onChange={(e) => onChange({ name: e.target.value })}
          className="h-7 w-40 font-mono text-[12px]"
          aria-label="Column name"
        />
        <Input
          value={field.label ?? ""}
          placeholder="Label"
          onChange={(e) => onChange({ label: e.target.value || undefined })}
          className="h-7 flex-1"
        />
        <SimpleSelect
          size="sm"
          className="w-32"
          value={field.type}
          onChange={(type) => onChange({ type: type as Field["type"], options: type === "option" ? (field.options ?? []) : undefined })}
          options={FIELD_TYPES.map((t) => ({ value: t, label: TYPE_LABELS[t] }))}
        />
        <IconButton
          label={field.required ? "Required" : "Optional"}
          className={cn(field.required ? "bg-muted text-destructive hover:text-destructive" : "opacity-50")}
          onClick={() => onChange({ required: !field.required || undefined })}
        >
          <AsteriskIcon />
        </IconButton>
        <div className="flex opacity-0 transition-opacity group-hover:opacity-100">
          <IconButton label="Move up" disabled={first} onClick={() => onMove(-1)}>
            <ArrowUpIcon />
          </IconButton>
          <IconButton label="Move down" disabled={last} onClick={() => onMove(1)}>
            <ArrowDownIcon />
          </IconButton>
          <IconButton label="Delete field" onClick={onRemove}>
            <Trash2Icon />
          </IconButton>
        </div>
      </div>
      <input
        value={field.description ?? ""}
        placeholder="Description (helps the agent)…"
        onChange={(e) => onChange({ description: e.target.value || undefined })}
        className="bg-transparent px-1 text-[12px] text-muted-foreground outline-none placeholder:text-muted-foreground/50"
      />
      {field.type === "option" && <OptionsEditor field={field} onChange={onChange} />}
    </div>
  );
}

function OptionsEditor({ field, onChange }: { field: Field; onChange: (patch: Partial<Field>) => void }) {
  const [draft, setDraft] = useState("");
  const options = field.options ?? [];
  const add = () => {
    const value = draft.trim();
    if (!value || options.some((o) => o.value === value)) return;
    onChange({ options: [...options, { value, color: COLORS[options.length % COLORS.length] }] });
    setDraft("");
  };
  return (
    <div className="flex flex-wrap items-center gap-1.5 px-1">
      {options.map((option, i) => (
        <span key={option.value} className={cn("inline-flex h-6 items-center gap-1 rounded-full pr-1 pl-1.5 text-[12px]", TINTS[option.color])}>
          <DropdownMenu>
            <DropdownMenuTrigger aria-label="Color" className={cn("size-2.5 rounded-full", DOTS[option.color])} />
            <DropdownMenuContent className="w-auto min-w-0">
              <div className="flex gap-1 p-1">
                {COLORS.map((color) => (
                  <DropdownMenuItem
                    key={color}
                    aria-label={color}
                    className="size-6 justify-center p-0"
                    onClick={() => onChange({ options: options.map((o, j) => (j === i ? { ...o, color } : o)) })}
                  >
                    <span className={cn("size-3 rounded-full", DOTS[color])} />
                  </DropdownMenuItem>
                ))}
              </div>
            </DropdownMenuContent>
          </DropdownMenu>
          {option.value}
          <button
            type="button"
            aria-label="Remove option"
            className="opacity-50 hover:opacity-100"
            onClick={() => onChange({ options: options.filter((_, j) => j !== i) })}
          >
            <XIcon className="size-3" />
          </button>
        </span>
      ))}
      <input
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && add()}
        onBlur={add}
        placeholder="Add option…"
        className="h-6 w-28 bg-transparent px-1 text-[12px] outline-none placeholder:text-muted-foreground/60"
      />
      <label className="ml-auto flex items-center gap-1.5 text-[12px] text-muted-foreground">
        Multiple
        <Switch size="sm" checked={!!field.multiple} onCheckedChange={(multiple) => onChange({ multiple: multiple || undefined })} />
      </label>
    </div>
  );
}
