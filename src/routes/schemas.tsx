import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import CodeMirror from "@uiw/react-codemirror";
import { json as jsonLanguage } from "@codemirror/lang-json";
import { z } from "zod";
import { ClipboardListIcon, DownloadIcon, InfoIcon, ExternalLinkIcon, FileSpreadsheetIcon, FolderIcon, LinkIcon, RefreshCwIcon, SaveIcon, SparklesIcon, UploadIcon, XIcon } from "lucide-react";
import { Chat } from "@/components/chat";
import { Badge } from "@/components/field-input";
import { Picker } from "@/components/picker";
import { Empty, PageHeader, SidePanel } from "@/components/side-panel";
import { AI_PANEL, EXCEL_UPLOAD } from "@/config";
import { tool, type Agent } from "@/agent/loop";
import { withContexts } from "@/agent/context";
import { applySchemaToExcel, detectFields, previewUrl, readRows, resolveFolder } from "@/lib/excel";
import { rowsText } from "@/lib/pbi";
import { applyForm, designUrl, planForm, type FormPlan } from "@/lib/forms";
import { downloadText, linkForm, saveFields } from "@/lib/library";
import { errorText } from "@/lib/ms";
import { COLORS, FIELD_TYPES, TYPE_LABELS, fieldsErrors, isIdField, type Field, type FormSection } from "@/lib/schema";
import { patch, useCollection, type SchemaEntry } from "@/lib/store";
import { useLast } from "@/lib/last";
import { fieldsText } from "./excel";

export const Route = createFileRoute("/schemas")({
  validateSearch: (s: Record<string, unknown>) => ({ schema: (s.schema as string) || undefined }),
  component: SchemaEditor,
});

type Tab = "json" | "excel";

/** What the JSON editor shows and takes back: the title, the fields and the form's sections (the files are set from the header). */
const editable = (s: SchemaEntry) => ({ title: s.title, fields: s.fields, sections: s.sections ?? [] });

/**
 * Schema editor: the definition of a schema's columns (types, choices with colors, rules), edited as
 * JSON or by the AI, then written to its Excel file (columns, dropdowns, colors) and to its Microsoft
 * Form (questions) with the header buttons.
 */
function SchemaEditor() {
  const { schema: schemaId } = Route.useSearch();
  const navigate = useNavigate({ from: "/schemas" });
  useLast("schemas:schema", schemaId, (schema) => navigate({ search: { schema }, replace: true }));
  const schemas = useCollection("schemas").filter((s) => !s.builtIn);
  const schema = schemas.find((s) => s.id === schemaId);
  const [tab, setTab] = useState<Tab>("json");
  const [state, setState] = useState<{ busy?: { label: string; done?: number; total?: number }; done?: string; error?: string }>({});
  const [plan, setPlan] = useState<FormPlan>();
  /** Bumped after the Excel file is written: the preview is loaded again. */
  const [version, setVersion] = useState(0);
  useEffect(() => (setState({}), setPlan(undefined)), [schemaId]);
  const run = (busy: string, task: (progress: (label: string, done: number, total: number) => void) => Promise<string | void>) => {
    setState({ busy: { label: busy } });
    task((label, done, total) => setState({ busy: { label, done, total } })).then((done) => setState({ done: done ?? undefined }), (e) => setState({ error: errorText(e) }));
  };
  const agent = useMemo(() => (schema ? schemaAgent(schema) : undefined), [schema]);

  const updateExcel = () =>
    confirm(`Write the schema to "${schema!.excel!.name}"?\n\nMissing columns are added${schema!.form ? " at the end (the columns the form writes to never move)" : ""} — none is deleted —, the hidden _schema sheet is rewritten, and the table's sheet gets the dropdowns and colors of the choices. The other sheets are not touched.`) &&
    run("Writing the Excel file…", async (progress) => {
      await applySchemaToExcel(schema!.excel!, schema!.fields, progress, { appendOnly: !!schema!.form });
      setVersion((v) => v + 1);
      return "Excel file updated.";
    });
  const reviewForm = () => run("Reading the form…", async () => setPlan(await planForm(schema!.form!, schema!.fields, schema!.sections)));
  /** Above the content: its tabs, and the form linked to the schema. */
  const toolbar = schema && (
    <>
      <button type="button" className="tab" data-active={tab === "json"} onClick={() => setTab("json")}>
        JSON
      </button>
      <button type="button" className="tab" data-active={tab === "excel"} onClick={() => setTab("excel")}>
        Excel
      </button>
      <span className="mx-1 h-6 shrink-0 border-l" />
      <FormLink schema={schema} />
      {EXCEL_UPLOAD && schema.fields.some((f) => f.type === "image" || f.type === "file") && <UploadsFolder schema={schema} />}
    </>
  );

  return (
    <div className="flex h-full">
      {AI_PANEL && <SidePanel tabs={[{ id: "ai", label: "AI assistant", icon: SparklesIcon, content: agent ? <Chat scope={`schemas:${schemaId}`} agent={agent} placeholder="Ask to add or change columns, choices, rules…" /> : <Empty>Choose a schema.</Empty> }]} />}
      <div className="flex min-w-0 flex-1 flex-col">
        <PageHeader
          screen="Schema editor"
          picker={
            <Picker
              items={schemas.map((s) => ({ id: s.id, label: s.title, hint: s.form ? "form" : undefined }))}
              value={schemaId}
              onChange={(schema) => navigate({ search: { schema } })}
              placeholder="Choose a schema…"
              empty={<>No schema: add one in the <Link to="/" search={{ tab: "schemas" }} className="underline">Library</Link>.</>}
            />
          }
          actions={
            schema && (
              <>
                <button type="button" className="btn h-9 rounded-full" disabled={!schema.form || !!state.busy} title={schema.form ? "Write the questions of the form after the schema (review first)" : "No form linked to this schema"} onClick={reviewForm}>
                  <UploadIcon /> Form
                </button>
                <button type="button" className="btn h-9 rounded-full" disabled={!!state.busy} title="Write the columns, dropdowns, colors and look to the Excel file" onClick={updateExcel}>
                  <FileSpreadsheetIcon /> Excel
                </button>
                <button type="button" className="icon-btn" title="Download the schema (JSON)" onClick={() => downloadText(`${schema.title}.schema.json`, JSON.stringify(schema, null, 2))}>
                  <DownloadIcon />
                </button>
                <a className="icon-btn" title="Open the Excel file" href={schema.excel!.webUrl ?? schema.excel!.link} target="_blank" rel="noreferrer">
                  <ExternalLinkIcon />
                </a>
                {schema.form && (
                  <a className="icon-btn" title="Open the form in Forms" href={designUrl(schema.form)} target="_blank" rel="noreferrer">
                    <ClipboardListIcon />
                  </a>
                )}
              </>
            )
          }
        >
          {state.busy && <Progress {...state.busy} />}
          {state.done && <span className="truncate text-[13px] text-green-700">{state.done}</span>}
          {state.error && <span className="max-w-md truncate text-[13px] text-destructive" title={state.error}>{state.error}</span>}
        </PageHeader>
        {!schema && <Empty>Choose a schema to edit its columns (built-in schemas of the report viewer cannot change).</Empty>}
        {schema && plan && (
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
        {schema && tab === "json" && <JsonEditor key={schema.id} schema={schema} toolbar={toolbar} />}
        {schema && tab === "excel" && <ExcelPreview key={`${schema.id}:${version}`} schema={schema} toolbar={toolbar} />}
      </div>
    </div>
  );
}

/** What a long write is doing, with a bar when its steps are counted. */
function Progress({ label, done, total }: { label: string; done?: number; total?: number }) {
  return (
    <span className="flex min-w-0 items-center gap-2 text-[13px] text-muted-foreground">
      {total ? (
        <span className="h-1.5 w-28 shrink-0 overflow-hidden rounded-full bg-muted">
          <span className="block h-full rounded-full bg-foreground transition-[width] duration-300" style={{ width: `${Math.round(((done ?? 0) / total) * 100)}%` }} />
        </span>
      ) : (
        <span className="size-1.5 shrink-0 animate-pulse rounded-full bg-foreground" />
      )}
      <span className="truncate">{label}</span>
      {total ? <span className="shrink-0 tabular-nums">{done}/{total}</span> : null}
    </span>
  );
}

/* ------------------------------- JSON editor ------------------------------- */

/** The schema as JSON: checked as it is typed (JSON, then fields), saved to the library on demand. */
function JsonEditor({ schema, toolbar }: { schema: SchemaEntry; toolbar: ReactNode }) {
  const saved = JSON.stringify(editable(schema), null, 2);
  const [text, setText] = useState(saved);
  const [state, setState] = useState<{ error?: string; done?: string }>({});
  // The AI (or another tab) changed the schema: shown at once when nothing is being typed.
  const last = useRef(saved);
  useEffect(() => {
    if (text === last.current) setText(saved);
    last.current = saved;
  }, [saved]); // eslint-disable-line react-hooks/exhaustive-deps
  const problems = useMemo(() => {
    try {
      const value = JSON.parse(text) as { title?: unknown; fields?: Field[]; sections?: FormSection[] };
      if (typeof value.title !== "string" || !value.title.trim()) return ['"title" must be a text.'];
      if (!Array.isArray(value.fields)) return ['"fields" must be a list.'];
      if (value.sections !== undefined && !Array.isArray(value.sections)) return ['"sections" must be a list.'];
      return fieldsErrors(value.fields, value.sections);
    } catch (e) {
      return [`Invalid JSON: ${errorText(e)}`];
    }
  }, [text]);
  const changed = text !== saved;
  const save = () => {
    const value = JSON.parse(text) as { title: string; fields: Field[]; sections?: FormSection[] };
    saveFields(schema, value.fields, value.sections ?? []).then(
      () => patch("schemas", schema.id, { title: value.title.trim() }).then(() => setState({ done: "Saved in the library. Write it to the files with the buttons above." })),
      (e) => setState({ error: errorText(e) }),
    );
  };
  return (
    <div className="flex min-h-0 flex-1 flex-col gap-3 p-6 pt-4">
      <div className="flex items-center gap-2">
        {toolbar}
        <span
          className="icon-btn cursor-help"
          title={`Field: name (= Excel column = question), type (${FIELD_TYPES.map((t) => `${t}: ${TYPE_LABELS[t]}`).join(", ")}), label, description, required, multiple, options [{value, color}] (colors: ${COLORS.join(", ")}), max, default, form (false: Excel only), branches {option: section | "end"}. Sections: [{title, fields, next}].`}
        >
          <InfoIcon />
        </span>
        <span className="ml-auto" />
        {changed && (
          <>
            <button type="button" className="btn" onClick={() => (setText(saved), setState({}))}>
              <XIcon /> Cancel
            </button>
            <button type="button" className="btn-primary" disabled={problems.length > 0} onClick={save}>
              <SaveIcon /> Save
            </button>
          </>
        )}
      </div>
      {problems.length > 0 && <p className="rounded-lg bg-red-500/5 px-3 py-2 text-[13px] text-destructive">{problems.join(" ")}</p>}
      {state.error && <p className="text-[13px] text-destructive">{state.error}</p>}
      {state.done && !changed && <p className="text-[13px] text-green-700">{state.done}</p>}
      <div className="card min-h-0 flex-1 overflow-hidden">
        <CodeMirror value={text} height="100%" className="h-full text-[13px]" extensions={[jsonLanguage()]} onChange={(v) => (setText(v), setState({}))} basicSetup={{ foldGutter: true, highlightActiveLine: false }} />
      </div>
    </div>
  );
}

/* ------------------------------ Excel preview ------------------------------ */

/** The Excel file as Excel for the web shows it (read only; the link opening it expires after a few minutes, so it is asked again). */
function ExcelPreview({ schema, toolbar }: { schema: SchemaEntry; toolbar: ReactNode }) {
  const [url, setUrl] = useState<string>();
  const [error, setError] = useState<string>();
  const load = () => {
    setUrl(undefined);
    previewUrl(schema.excel!).then(setUrl, (e) => setError(errorText(e)));
  };
  useEffect(load, [schema.excel?.itemId]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div className="flex min-h-0 flex-1 flex-col gap-3 p-6 pt-4">
      <div className="flex items-center gap-2 text-[13px] text-muted-foreground">
        {toolbar}
        <span className="ml-auto truncate">Read only · {schema.excel!.name} › {schema.excel!.table}</span>
        <button type="button" className="icon-btn" title="Load again" onClick={load}>
          <RefreshCwIcon />
        </button>
      </div>
      {error && <p className="text-[13px] text-destructive">{error}</p>}
      <div className="card min-h-0 flex-1 overflow-hidden">{url ? <iframe title="Excel" src={url} className="size-full border-0" /> : !error && <Empty>Loading…</Empty>}</div>
    </div>
  );
}

/* ----------------------------------- Form ---------------------------------- */

/** The form linked to the schema: its name, or a field to link one (its questions can retype the matching columns). */
function FormLink({ schema }: { schema: SchemaEntry }) {
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  if (schema.form)
    return (
      <span className="flex items-center gap-1.5">
        <Badge color="violet">{schema.form.title || "Form"}</Badge>
        <button type="button" className="icon-btn" title="Unlink the form (the form is kept)" onClick={() => confirm("Unlink the form from this schema? (the form itself is kept)") && void patch("schemas", schema.id, { form: undefined })}>
          <XIcon />
        </button>
      </span>
    );
  return (
    <span className="flex items-center gap-1.5">
      <input className="input h-9 w-52 rounded-full text-[13px]" placeholder="Form link (optional)" value={link} onChange={(e) => setLink(e.target.value)} />
      <button
        type="button"
        className="icon-btn"
        title="Link this form"
        disabled={!link.trim() || state.busy}
        onClick={() => {
          setState({ busy: true });
          linkForm(schema, link.trim(), confirm("Give the matching columns the types and choices of the form's questions?")).then(() => (setState({}), setLink("")), (e) => setState({ error: errorText(e) }));
        }}
      >
        <LinkIcon className={state.busy ? "animate-pulse" : ""} />
      </button>
      {state.error && <span className="text-[13px] text-destructive">{state.error}</span>}
    </span>
  );
}

/**
 * Optional: the SharePoint folder receiving the files uploaded in the Excel editor (image and file
 * fields); without it those fields only take links.
 */
function UploadsFolder({ schema }: { schema: SchemaEntry }) {
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  if (schema.filesFolder)
    return (
      <span className="flex items-center gap-1.5" title="Files uploaded in the Excel editor go to this folder; their link is stored in the field">
        <a href={schema.filesFolder.link} target="_blank" rel="noreferrer">
          <Badge color="blue">{`Uploads: ${schema.filesFolder.name}`}</Badge>
        </a>
        <button type="button" className="icon-btn" title="No uploads (links only)" onClick={() => void patch("schemas", schema.id, { filesFolder: undefined })}>
          <XIcon />
        </button>
      </span>
    );
  return (
    <span className="flex items-center gap-1.5">
      <input className="input h-9 w-56 rounded-full text-[13px]" placeholder="Uploads folder link (optional)" title="SharePoint folder receiving the files uploaded in the Excel editor (image and file fields)" value={link} onChange={(e) => setLink(e.target.value)} />
      <button
        type="button"
        className="icon-btn"
        title="Use this folder for uploads"
        disabled={!link.trim() || state.busy}
        onClick={() => {
          setState({ busy: true });
          resolveFolder(link.trim())
            .then((filesFolder) => patch("schemas", schema.id, { filesFolder }))
            .then(() => (setState({}), setLink("")), (e) => setState({ error: errorText(e) }));
        }}
      >
        <FolderIcon className={state.busy ? "animate-pulse" : ""} />
      </button>
      {state.error && <span className="text-[13px] text-destructive">{state.error}</span>}
    </span>
  );
}

/** What updating the form will do, to confirm (deleting a question deletes its answers in the form). */
function FormReview({ plan, onApply, onCancel }: { plan: FormPlan; onApply: () => void; onCancel: () => void }) {
  const nothing = !plan.create.length && !plan.update.length && !plan.remove.length;
  const list = (title: string, items: string[], color: "green" | "blue" | "red") =>
    items.length > 0 && (
      <div className="flex flex-col gap-1.5">
        <span className="font-medium">{title}</span>
        <span className="flex flex-wrap gap-1.5">
          {items.map((t, i) => (
            <Badge key={i} color={color}>{t}</Badge>
          ))}
        </span>
      </div>
    );
  return (
    <div className="mx-6 mt-4 flex flex-col gap-3 rounded-xl border border-amber-300 bg-amber-50 p-4">
      <span className="font-semibold">{nothing ? "The form already matches the schema." : "Update the form?"}</span>
      {list("To add", plan.create, "green")}
      {list("To change (title, type, choices, branches, order)", plan.update, "blue")}
      {list("Questions to delete (their answers are deleted from the form)", plan.remove.map((q) => q.title), "red")}
      <div className="flex gap-2">
        {!nothing && (
          <button type="button" className="btn-primary" onClick={onApply}>
            Update the form
          </button>
        )}
        <button type="button" className="btn" onClick={onCancel}>
          {nothing ? "Close" : "Cancel"}
        </button>
      </div>
    </div>
  );
}

/* ----------------------------------- AI ------------------------------------ */

const option = z.object({ value: z.string(), color: z.enum(COLORS).optional() });
const fieldArgs = {
  type: z.enum(FIELD_TYPES).optional(),
  label: z.string().optional().describe("Shown in the app (the form's question is the column name)"),
  description: z.string().optional().describe("Help text (also the question's subtitle in the form)"),
  required: z.boolean().optional(),
  multiple: z.boolean().optional().describe("choice: several values allowed"),
  options: z.array(option).optional().describe("choice: the values, in order, each with a color"),
  max: z.number().optional().describe("rating: 5 or 10"),
  in_form: z.boolean().optional().describe("false: an Excel column only, never asked in the form"),
  default: z.union([z.string(), z.number(), z.array(z.string())]).nullable().optional().describe("Value of new rows (null: none)"),
  branches: z.record(z.string(), z.string()).optional().describe('single choice: option → section title or "end" (where the form goes after that answer); {} removes them'),
};
type FieldArgs = { [K in keyof typeof fieldArgs]?: z.infer<(typeof fieldArgs)[K]> };

/** Fields and sections before the AI's last save, per schema (the agent is made again when the schema changes). */
const previousSchemas = new Map<string, { fields: Field[]; sections: FormSection[] }>();

/**
 * The AI edits a draft of the fields and sections (several changes, checked together) saved to the
 * library when its answer is done; the previous ones are kept to undo the last save.
 */
function schemaAgent(schema: SchemaEntry): Agent {
  let draft: { fields: Field[]; sections: FormSection[] } | undefined;
  const current = () => (draft ??= structuredClone({ fields: schema.fields, sections: schema.sections ?? [] }));
  const fields = () => current().fields;
  const find = (name: string) => {
    const f = fields().find((x) => x.name.toLowerCase() === name.toLowerCase());
    if (!f) throw new Error(`No field "${name}". Fields: ${fields().map((x) => x.name).join(", ")}.`);
    return f;
  };
  const colored = (options?: { value: string; color?: (typeof COLORS)[number] }[]) => options?.map((o, i) => ({ value: o.value, color: o.color ?? COLORS[(i + 1) % COLORS.length] }));
  /** A field with the given settings (only those given). */
  const withArgs = (f: Field, a: FieldArgs): Field => {
    const next: Field = { ...f, ...Object.fromEntries(Object.entries({ type: a.type, label: a.label, description: a.description, required: a.required, multiple: a.multiple, max: a.max }).filter(([, v]) => v !== undefined)) };
    if (a.options) next.options = colored(a.options);
    if (next.type !== "option") delete next.options;
    if (a.default !== undefined) {
      if (a.default === null) delete next.default;
      else next.default = a.default;
    }
    if (a.in_form !== undefined) {
      if (a.in_form) delete next.form;
      else next.form = false;
    }
    if (a.branches) {
      if (Object.keys(a.branches).length) next.branches = a.branches;
      else delete next.branches;
    }
    return next;
  };
  const checked = (next: Partial<NonNullable<typeof draft>>) => {
    const state = { ...current(), ...next };
    const errors = fieldsErrors(state.fields, state.sections);
    if (errors.length) throw new Error(`${errors.join(" ")} Nothing was changed.`);
    draft = state;
  };
  const sectionsText = (list: FormSection[]) =>
    list.length ? list.map((x) => `- section "${x.title}": ${x.fields.join(", ") || "no question"}${x.next ? ` → then ${x.next === "end" ? "END of the form" : `"${x.next}"`}` : ""}`).join("\n") : "No section (one page).";
  return {
    instructions: `You edit the SCHEMA of an Excel table and of the Microsoft Form filling it (when linked): its columns, their types, choices with colors, rules, and the form's sections and branches. You never change the data (that is the Excel editor's job).
Types (the ones a Microsoft Form can ask): ${FIELD_TYPES.map((t) => `${t} = ${TYPE_LABELS[t]}`).join("; ")}. Colors of choices: ${COLORS.join(", ")}.
- A field's "name" is its Excel column AND its question in the form (Forms names the Excel columns after the questions): write it as the question should read; never rename an existing one (a new name is a new, empty column). "label" only changes the app's display. The "id" column is the app's: leave it.
- in_form false: a column of the Excel only (e.g. a status set by the team), never asked.
- Sections (set_sections): groups of questions, each with where the form goes when done (a section, or "end"). Questions in no section come first. Branches (update_field branches): on a single choice, where the form goes after each option (a section, or "end").
Your changes are saved to the library when your answer is done; the user then writes them with the buttons "Update the Excel" and "Update the form" (say so when you changed something).`,
    context: async () => {
      // What the table holds now (first rows, every column of the file): to fit the schema to the data.
      const rows = schema.excel
        ? await readRows(schema.excel, (await detectFields(schema.excel)).map((f) => ({ ...f, type: "string" as const })))
            .then((t) => `The Excel table now (${t.rows.length} rows; first ${Math.min(15, t.rows.length)}, every column of the file):\n${rowsText(t.rows.slice(0, 15).map((r) => Object.fromEntries(t.header.map((h) => [h, r[h]]))), 15)}`)
            .catch((e) => `The Excel table could not be read: ${e instanceof Error ? e.message : e}`)
        : "";
      return withContexts([`Schema "${schema.title}": Excel table "${schema.excel?.table}" of "${schema.excel?.name}"${schema.form ? `; Microsoft Form "${schema.form.title}"` : "; no form linked"}.`, rows], [schema.contextIds]);
    },
    state: async () => {
      const { fields: list, sections } = current();
      const extra = (f: Field) => [f.default !== undefined && `default ${JSON.stringify(f.default)}`, f.form === false && "not in the form", f.branches && `branches: ${Object.entries(f.branches).map(([o, t]) => `${o} → ${t}`).join(", ")}`].filter(Boolean).join("; ");
      return `Fields now${draft ? " (with your unsaved changes)" : ""}:\n${fieldsText(list).split("\n").map((line, i) => (extra(list[i]) ? `${line} [${extra(list[i])}]` : line)).join("\n")}\n\nForm sections:\n${sectionsText(sections)}`;
    },
    tools: [
      tool({
        name: "add_field",
        description: "Adds a field (a new column at the end, or after another field; a new question when the form is updated).",
        args: z.object({ ...fieldArgs, name: z.string(), type: z.enum(FIELD_TYPES), after: z.string().optional() }),
        run: async ({ name, type, after, ...rest }) => {
          if (fields().some((f) => f.name.toLowerCase() === name.toLowerCase())) throw new Error(`"${name}" already exists.`);
          const list = [...fields()];
          list.splice(after ? list.indexOf(find(after)) + 1 : list.length, 0, withArgs({ name, type }, { ...rest, type }));
          checked({ fields: list });
          return `Field "${name}" added (${TYPE_LABELS[type]}).`;
        },
      }),
      tool({
        name: "update_field",
        description: "Changes a field (only the given settings; options replace the whole list of choices).",
        args: z.object({ name: z.string(), ...fieldArgs }),
        run: async ({ name, ...rest }) => {
          const f = find(name);
          if (isIdField(f)) throw new Error("The id column is the app's: it cannot change.");
          checked({ fields: fields().map((x) => (x === f ? withArgs(f, rest) : x)) });
          return `Field "${f.name}" changed.`;
        },
      }),
      tool({
        name: "delete_field",
        description: "Removes a field from the schema (the Excel column and its data are kept, only hidden from the app; its question goes when the form is updated).",
        args: z.object({ name: z.string() }),
        run: async ({ name }) => {
          const f = find(name);
          if (isIdField(f)) throw new Error("The id column cannot be removed.");
          const sections = current().sections.map((x) => ({ ...x, fields: x.fields.filter((n) => n !== f.name) }));
          checked({ fields: fields().filter((x) => x !== f), sections });
          return `Field "${f.name}" removed.`;
        },
      }),
      tool({
        name: "reorder_fields",
        description: "Sets the order of the fields (every field name, in the new order): the Excel columns and the questions follow it.",
        args: z.object({ names: z.array(z.string()) }),
        run: async ({ names }) => {
          const list = names.map(find);
          const missing = fields().filter((f) => !list.includes(f));
          if (missing.length) throw new Error(`Missing: ${missing.map((f) => f.name).join(", ")}. Give every field.`);
          checked({ fields: list });
          return "Order changed.";
        },
      }),
      tool({
        name: "set_sections",
        description: 'Replaces the form\'s sections (all of them, in order): title, description, its questions (field names, in order), next (a section title, or "end"; default: the next section). [] removes them.',
        args: z.object({ sections: z.array(z.object({ title: z.string(), description: z.string().optional(), fields: z.array(z.string()), next: z.string().optional() })) }),
        run: async ({ sections }) => {
          const old = current().sections;
          // A section keeps its form question when its title stays.
          checked({ sections: sections.map((x) => ({ ...x, fields: x.fields.map((n) => find(n).name), question: old.find((o) => o.title === x.title)?.question })) });
          return `${sections.length} section(s) set.`;
        },
      }),
      tool({
        name: "discard_changes",
        description: "Drops your unsaved changes.",
        args: z.object({}),
        run: async () => ((draft = undefined), "Changes dropped."),
      }),
      tool({
        name: "undo_last_save",
        description: "Puts back the fields and sections as they were before your last saved changes (only when the user asks).",
        args: z.object({}),
        run: async () => {
          const previous = previousSchemas.get(schema.id);
          if (!previous) throw new Error("Nothing to undo in this session.");
          await saveFields(schema, previous.fields, previous.sections);
          previousSchemas.delete(schema.id);
          draft = undefined;
          return "The previous fields and sections are back.";
        },
      }),
    ],
    finish: async () => {
      if (!draft || JSON.stringify(draft) === JSON.stringify({ fields: schema.fields, sections: schema.sections ?? [] })) return (draft = undefined);
      previousSchemas.set(schema.id, { fields: schema.fields, sections: schema.sections ?? [] });
      await saveFields(schema, draft.fields, draft.sections);
      draft = undefined;
      return "✓ Schema saved in the library (use “Update the Excel” / “Update the form” to write it to the files).";
    },
  };
}
