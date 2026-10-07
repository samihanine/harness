import { z } from "zod";
import { normalizeRich } from "./rich";

export const FIELD_TYPES = ["string", "text", "number", "integer", "boolean", "date", "option", "image"] as const;
export type FieldType = (typeof FIELD_TYPES)[number];

export const COLORS = ["gray", "red", "orange", "amber", "green", "teal", "blue", "violet", "pink"] as const;

const option = z.object({ value: z.string().min(1), color: z.enum(COLORS).default("gray") });

export const fieldSchema = z.object({
  name: z
    .string()
    .min(1)
    .regex(/^[A-Za-z_][\w]*$/, "letters, digits and _ only")
    .refine((n) => n !== "id", "id is reserved"),
  type: z.enum(FIELD_TYPES),
  label: z.string().optional(),
  description: z.string().optional(),
  required: z.boolean().optional(),
  /** option: several values per cell (stored "a; b"). */
  multiple: z.boolean().optional(),
  options: z.array(option).optional(),
});

export const schemaSchema = z
  .object({
    id: z.string(),
    name: z.string().min(1),
    /** Worksheet holding the rows (header on line 1). */
    sheet: z.string().min(1).max(31),
    description: z.string().optional(),
    fields: z.array(fieldSchema),
  })
  .refine((s) => new Set(s.fields.map((f) => f.name)).size === s.fields.length, "field names must be unique");

export type Field = z.infer<typeof fieldSchema>;
export type OptionValue = z.infer<typeof option>;
export type Schema = z.infer<typeof schemaSchema>;
export type Row = { id: string; [field: string]: unknown };

export const labelOf = (field: Field) => {
  const label = field.label || field.name.replace(/_/g, " ");
  return label.charAt(0).toUpperCase() + label.slice(1);
};

export const newSchema = (name: string): Schema => ({
  id: crypto.randomUUID(),
  name,
  sheet: name.slice(0, 31) || "Data",
  fields: [{ name: "title", type: "string", label: "Title", required: true }],
});

/** Empty value of a field. */
export const emptyValue = (field: Field) =>
  field.type === "boolean" ? false : field.type === "option" && field.multiple ? [] : null;

/** Coerces a value (from the UI, the agent or a cell) to the field type; throws on invalid values. */
export function coerce(field: Field, value: unknown): unknown {
  if (value === undefined || value === null || value === "") return emptyValue(field);
  switch (field.type) {
    case "number":
    case "integer": {
      const n = typeof value === "number" ? value : Number(String(value).replace(",", "."));
      if (Number.isNaN(n)) throw new Error(`${field.name}: "${value}" is not a number`);
      return field.type === "integer" ? Math.round(n) : n;
    }
    case "boolean":
      return value === true || /^(true|yes|oui|1|x)$/i.test(String(value));
    case "date": {
      const date = value instanceof Date ? value : new Date(String(value));
      if (Number.isNaN(+date)) throw new Error(`${field.name}: "${value}" is not a date`);
      return date.toISOString().slice(0, 10);
    }
    case "option": {
      const values = (Array.isArray(value) ? value : String(value).split(/\s*;\s*/)).map(String).filter(Boolean);
      const allowed = new Set(field.options?.map((o) => o.value));
      const unknown = values.filter((v) => !allowed.has(v));
      if (unknown.length) throw new Error(`${field.name}: unknown option(s) ${unknown.join(", ")} (allowed: ${[...allowed].join(", ")})`);
      return field.multiple ? values : (values[0] ?? null);
    }
    case "text":
      return normalizeRich(value);
    default:
      return String(value);
  }
}
