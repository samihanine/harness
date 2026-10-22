/**
 * The Excel screen's schema tools: the AI edits a draft of a schema's fields (columns: types, choices
 * with colors, rules) and of its form's sections, checked together and saved when its answer is done.
 * Writing them to the Excel file and the form stays a button of the Schema tab (the user reviews it).
 */
import { z } from "zod";
import { tool, type Agent } from "./loop";
import { withContexts } from "./context";
import { fieldsText } from "./excel-tools";
import { detectFields, readRows } from "@/lib/excel";
import { saveFields } from "@/lib/library";
import { rowsText } from "@/lib/pbi";
import { COLORS, FIELD_TYPES, TYPE_LABELS, fieldsErrors, isIdField, type Field, type FormSection } from "@/lib/schema";
import type { SchemaEntry } from "@/lib/store";

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
 * app when its answer is done; the previous ones are kept to undo the last save.
 */
export function schemaAgent(schema: SchemaEntry): Agent {
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
    instructions: `SCHEMA (Schema tab): you edit the columns of the Excel table and of the Microsoft Form filling it (when linked): their types, choices with colors, rules, and the form's sections and branches.
Types (the ones a Microsoft Form can ask): ${FIELD_TYPES.map((t) => `${t} = ${TYPE_LABELS[t]}`).join("; ")}. Colors of choices: ${COLORS.join(", ")}.
- A field's "name" is its Excel column AND its question in the form (Forms names the Excel columns after the questions): write it as the question should read; never rename an existing one (a new name is a new, empty column). "label" only changes the app's display. The "id" column is the app's: leave it.
- in_form false: a column of the Excel only (e.g. a status set by the team), never asked.
- Sections (set_sections): groups of questions, each with where the form goes when done (a section, or "end"). Questions in no section come first. Branches (update_field branches): on a single choice, where the form goes after each option (a section, or "end").
Schema changes are saved in the app when your answer is done; the user then writes them to the files with the buttons "Sync the Excel" and "Update the form" of the Schema tab (say so when you changed the schema).`,
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
      return "✓ Schema saved in the app.";
    },
  };
}
