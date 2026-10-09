/**
 * Fields of an Excel table. The schema lives in the workbook itself (hidden sheet "_schema"), so the
 * file is self-describing: whoever opens it in the app gets the same option colors, labels, types.
 */
export const FIELD_TYPES = ["string", "text", "number", "integer", "boolean", "date", "option", "image"] as const;
export type FieldType = (typeof FIELD_TYPES)[number];
export const TYPE_LABELS: Record<FieldType, string> = {
  string: "Short text",
  text: "Rich text (markdown)",
  number: "Number",
  integer: "Integer",
  boolean: "Yes / no",
  date: "Date",
  option: "Option",
  image: "Image (URL)",
};

export const COLORS = ["gray", "red", "orange", "amber", "green", "teal", "blue", "violet", "pink"] as const;
export type Color = (typeof COLORS)[number];
export type Option = { value: string; color: Color };

export type Field = {
  name: string;
  type: FieldType;
  label?: string;
  description?: string;
  required?: boolean;
  /** option: several values per cell (stored "a; b"). */
  multiple?: boolean;
  options?: Option[];
};

export type Row = Record<string, unknown>;

/** Technical columns (row id) are kept in the file but not shown. */
export const isHidden = (f: Field) => /^(id|_id|row_?id)$/i.test(f.name);

/** Field used as the title of a row: title / name / label…, else the first short text. */
export const titleField = (fields: Field[]) =>
  fields.find((f) => !isHidden(f) && /^(title|titre|name|nom|label|libell[ée])$/i.test(f.name)) ?? fields.find((f) => !isHidden(f) && f.type === "string") ?? fields.find((f) => !isHidden(f));

export const labelOf = (f: Field) => {
  const label = f.label || f.name.replace(/_/g, " ");
  return label.charAt(0).toUpperCase() + label.slice(1);
};

/** Tinted badge per option color. */
export const TINTS: Record<Color, string> = {
  gray: "bg-zinc-500/10 text-zinc-700",
  red: "bg-red-500/10 text-red-700",
  orange: "bg-orange-500/10 text-orange-700",
  amber: "bg-amber-500/15 text-amber-800",
  green: "bg-emerald-500/10 text-emerald-700",
  teal: "bg-teal-500/10 text-teal-700",
  blue: "bg-blue-500/10 text-blue-700",
  violet: "bg-violet-500/10 text-violet-700",
  pink: "bg-pink-500/10 text-pink-700",
};

/** Solid swatch per color (color pickers). */
export const SWATCHES: Record<Color, string> = {
  gray: "bg-zinc-400",
  red: "bg-red-500",
  orange: "bg-orange-500",
  amber: "bg-amber-500",
  green: "bg-emerald-500",
  teal: "bg-teal-500",
  blue: "bg-blue-500",
  violet: "bg-violet-500",
  pink: "bg-pink-500",
};
export const isColor = (v: unknown): v is Color => (COLORS as readonly unknown[]).includes(v);

/** Values of a cell as a list (options). */
export const listOf = (value: unknown) => (Array.isArray(value) ? value : value === null || value === undefined || value === "" ? [] : String(value).split(/\s*;\s*/)).map(String).filter(Boolean);

/** Coerces a value (UI, AI or cell) to the field type; throws on invalid values. */
export function coerce(field: Field, value: unknown): unknown {
  if (value === undefined || value === null || value === "") return field.type === "boolean" ? false : field.type === "option" && field.multiple ? [] : null;
  switch (field.type) {
    case "number":
    case "integer": {
      const n = typeof value === "number" ? value : Number(String(value).replace(",", "."));
      if (Number.isNaN(n)) throw new Error(`${field.name}: "${value}" is not a number`);
      return field.type === "integer" ? Math.round(n) : n;
    }
    case "boolean":
      return value === true || /^(true|yes|oui|vrai|1|x)$/i.test(String(value));
    case "date": {
      const date = typeof value === "number" ? new Date(Math.round((value - 25569) * 86400000)) : new Date(String(value));
      if (Number.isNaN(+date)) throw new Error(`${field.name}: "${value}" is not a date`);
      return date.toISOString().slice(0, 10);
    }
    case "option": {
      const values = listOf(value);
      const allowed = new Set(field.options?.map((o) => o.value));
      const unknown = values.filter((v) => !allowed.has(v));
      if (unknown.length) throw new Error(`${field.name}: unknown option(s) ${unknown.join(", ")} (allowed: ${[...allowed].join(", ")})`);
      return field.multiple ? values : (values[0] ?? null);
    }
    default:
      return String(value);
  }
}

/** Value as an Excel cell: lists joined with "; ", empty as "". */
export const toCell = (value: unknown) => (Array.isArray(value) ? value.join("; ") : value === null || value === undefined ? "" : value);

/** Schema of a table without "_schema" sheet, guessed from its header, cell types and values. */
export function inferFields(header: string[], types: string[][], formats: string[][], values: unknown[][]): Field[] {
  return header.map((name, i) => {
    const column = values.map((r) => r[i]).filter((v) => v !== "" && v !== null);
    const kinds = types.map((t) => t[i]).filter((t) => t !== "Empty");
    const format = String(formats[0]?.[i] ?? "");
    let type: FieldType = "string";
    if (kinds[0] === "Boolean") type = "boolean";
    else if (kinds[0] === "Double") type = /[dy]/i.test(format) && !/[#0]/.test(format.replace(/\[.*?\]/g, "")) ? "date" : column.every((v) => Number.isInteger(v)) ? "integer" : "number";
    else if (/image|photo|picture|logo/i.test(name) || (column.length && column.every((v) => /^https?:\/\/\S+\.(png|jpe?g|gif|webp|svg)(\?|$)/i.test(String(v))))) type = "image";
    else if (column.some((v) => String(v).length > 80 || /\n|\*\*|^#/.test(String(v)))) type = "text";
    else {
      const distinct = [...new Set(column.flatMap((v) => listOf(v)))];
      if (column.length >= 3 && distinct.length > 0 && distinct.length <= Math.max(3, column.length / 2) && distinct.length <= 12)
        return { name, type: "option", multiple: column.some((v) => String(v).includes(";")), options: distinct.map((value, k) => ({ value, color: COLORS[(k + 1) % COLORS.length] })) };
    }
    return { name, type };
  });
}

/* "_schema" sheet: one row per field. */
export const SCHEMA_HEADER = ["name", "type", "label", "description", "required", "multiple", "options"] as const;
export const fieldToRow = (f: Field) => [f.name, f.type, f.label ?? "", f.description ?? "", !!f.required, !!f.multiple, (f.options ?? []).map((o) => `${o.value}:${o.color}`).join("; ")];
export const rowToField = (r: unknown[]): Field => ({
  name: String(r[0]),
  type: (FIELD_TYPES as readonly string[]).includes(String(r[1])) ? (r[1] as FieldType) : "string",
  label: String(r[2] ?? "") || undefined,
  description: String(r[3] ?? "") || undefined,
  required: r[4] === true || undefined,
  multiple: r[5] === true || undefined,
  options: String(r[6] ?? "")
    .split(/\s*;\s*/)
    .filter(Boolean)
    .map((o) => {
      const [value, color] = o.split(/:(?=[a-z]+$)/);
      return { value, color: (COLORS as readonly string[]).includes(color) ? (color as Color) : "gray" };
    }),
});
