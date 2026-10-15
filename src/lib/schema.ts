/**
 * Fields of a schema (an Excel table, and optionally the Microsoft Form filling it). The types are the
 * ones a Microsoft Form can ask, so a schema can always be written to its form: short / long text,
 * number, date, choice, rating, link (a short text in the form), image and file (a file upload). The `id` column is the app's (never in the form).
 */
export const FIELD_TYPES = ["string", "text", "number", "date", "option", "rating", "url", "image", "file"] as const;
export type FieldType = (typeof FIELD_TYPES)[number];
export const TYPE_LABELS: Record<FieldType, string> = {
  string: "Short text",
  text: "Long text (markdown)",
  number: "Number",
  date: "Date",
  option: "Choice",
  rating: "Rating",
  url: "Link (URL)",
  image: "Image (link)",
  file: "File (link)",
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
  /** rating: highest value (5 or 10). */
  max?: number;
  /** Id of the question of the linked Microsoft Form (kept when the field is renamed). */
  question?: string;
  /** Value of new rows (created in the app) when none is given. */
  default?: string | number | string[];
  /** false: an Excel column only, never asked in the form (e.g. a status set by the team). */
  form?: false;
  /** choice: where the form goes after each option (a section's title, or "end"); others: the next question. */
  branches?: Record<string, string>;
};

/**
 * A section of the form: its questions (field names, in order) and where it goes when done (another
 * section's title, or "end"; default: the next section). Fields in no section come before the first one.
 */
export type FormSection = { title: string; description?: string; fields: string[]; next?: string; question?: string };

export type Row = Record<string, unknown>;

/* ---------------------------------- Row ids --------------------------------- */

/** Row id column (shown, read only): values like TASK_001, given by the app only. */
export const isIdField = (f: Pick<Field, "name">) => /^(id|_id|row_?id)$/i.test(f.name);
export const ID_FIELD: Field = { name: "id", type: "string", label: "ID", description: "Row id, given by the app (e.g. TASK_001)" };

/** Id prefix of a table: its name in capitals, singular ("Tasks" → TASK, "Project list" → PROJECT_LIST). */
export function idPrefix(table: string) {
  const name = table.toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "ROW";
  return name.length > 3 && /[^S]S$/.test(name) ? name.slice(0, -1) : name;
}

/**
 * The only place ids are made: `count` new ids after the highest one of this prefix already used
 * (other values, e.g. older UUIDs, are kept as they are and ignored). PREFIX_001, PREFIX_002…
 */
export function nextIds(prefix: string, existing: unknown[], count: number) {
  const pattern = new RegExp(`^${prefix}_(\\d+)$`);
  const max = Math.max(0, ...existing.map((v) => Number(String(v ?? "").trim().match(pattern)?.[1] ?? 0)));
  return Array.from({ length: count }, (_, i) => `${prefix}_${String(max + 1 + i).padStart(3, "0")}`);
}

/** Field used as the title of a row: title / name / label…, else the first short text. */
export const titleField = (fields: Field[]) =>
  fields.find((f) => !isIdField(f) && /^(title|titre|name|nom|label|libell[ée]|subject|sujet)$/i.test(f.name)) ?? fields.find((f) => !isIdField(f) && f.type === "string") ?? fields.find((f) => !isIdField(f));

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

/** Solid swatch per color (color pickers, dots). */
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
  if (value === undefined || value === null || value === "") return field.type === "option" && field.multiple ? [] : null;
  switch (field.type) {
    case "number": {
      const n = typeof value === "number" ? value : Number(String(value).replace(",", "."));
      if (Number.isNaN(n)) throw new Error(`${field.name}: "${value}" is not a number`);
      return n;
    }
    case "rating": {
      const n = Math.round(Number(value));
      const max = field.max ?? 5;
      if (Number.isNaN(n) || n < 1 || n > max) throw new Error(`${field.name}: a rating is a whole number from 1 to ${max} ("${value}")`);
      return n;
    }
    case "date": {
      const date = typeof value === "number" ? new Date(Math.round((value - 25569) * 86400000)) : new Date(String(value));
      if (Number.isNaN(+date)) throw new Error(`${field.name}: "${value}" is not a date`);
      return date.toISOString().slice(0, 10);
    }
    case "option": {
      // Yes / no cells of older files (Excel booleans) are the choices Yes and No.
      const values = typeof value === "boolean" ? [value ? "Yes" : "No"] : listOf(value);
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

const IMAGE_URL = /^https?:\/\/\S+\.(png|jpe?g|gif|webp|svg)(\?|$)/i;
export const isImageUrl = (v: unknown) => IMAGE_URL.test(String(v ?? ""));

/** Schema of a table without schema, guessed from its header, cell types and values. */
export function inferFields(header: string[], types: string[][], formats: string[][], values: unknown[][]): Field[] {
  return header.map((name, i) => {
    if (isIdField({ name })) return { ...ID_FIELD, name };
    const column = values.map((r) => r[i]).filter((v) => v !== "" && v !== null);
    const kinds = types.map((t) => t[i]).filter((t) => t !== "Empty");
    const format = String(formats[0]?.[i] ?? "");
    let type: FieldType = "string";
    if (kinds[0] === "Boolean") return { name, type: "option", options: [{ value: "Yes", color: "green" }, { value: "No", color: "gray" }] };
    if (kinds[0] === "Double") type = /[dy]/i.test(format) && !/[#0]/.test(format.replace(/\[.*?\]/g, "")) ? "date" : "number";
    else if (/image|photo|picture|logo/i.test(name) || (column.length && column.every(isImageUrl))) type = "image";
    else if (/file|fichier|attachment/i.test(name)) type = "file";
    else if (/url|link|lien/i.test(name) || (column.length && column.every((v) => /^https?:\/\/\S+$/i.test(String(v))))) type = "url";
    else if (column.some((v) => String(v).length > 80 || /\n|\*\*|^#/.test(String(v)))) type = "text";
    else {
      const distinct = [...new Set(column.flatMap((v) => listOf(v)))];
      if (column.length >= 3 && distinct.length > 0 && distinct.length <= Math.max(3, column.length / 2) && distinct.length <= 12)
        return { name, type: "option", multiple: column.some((v) => String(v).includes(";")) || undefined, options: distinct.map((value, k) => ({ value, color: COLORS[(k + 1) % COLORS.length] })) };
    }
    return { name, type };
  });
}

/* "_schema" sheet of the workbook: one row per field (the file describes itself). */
export const SCHEMA_HEADER = ["name", "type", "label", "description", "required", "multiple", "options", "max"] as const;
export const fieldToRow = (f: Field) => [f.name, f.type, f.label ?? "", f.description ?? "", !!f.required, !!f.multiple, (f.options ?? []).map((o) => `${o.value}:${o.color}`).join("; "), f.max ?? ""];
/** Older type names (integer, boolean) are read as their current ones. */
const OLD_TYPES: Record<string, FieldType> = { integer: "number", boolean: "option" };
export const rowToField = (r: unknown[]): Field => {
  const raw = String(r[1] ?? "");
  const type = (FIELD_TYPES as readonly string[]).includes(raw) ? (raw as FieldType) : (OLD_TYPES[raw] ?? "string");
  const options = String(r[6] ?? "")
    .split(/\s*;\s*/)
    .filter(Boolean)
    .map((o) => {
      const [value, color] = o.split(/:(?=[a-z]+$)/);
      return { value, color: isColor(color) ? color : ("gray" as Color) };
    });
  return {
    name: String(r[0]),
    type,
    label: String(r[2] ?? "") || undefined,
    description: String(r[3] ?? "") || undefined,
    required: r[4] === true || undefined,
    multiple: r[5] === true || undefined,
    options: type === "option" ? (raw === "boolean" && !options.length ? [{ value: "Yes", color: "green" }, { value: "No", color: "gray" }] : options) : undefined,
    max: Number(r[7]) || undefined,
  };
};

/** Checks a list of fields (names, duplicates, options, sections, branches): the reasons it cannot be saved, empty when fine. */
export function fieldsErrors(fields: Field[], sections: FormSection[] = []) {
  const errors: string[] = [];
  const names = new Set<string>();
  for (const f of fields) {
    if (!/^[\p{L}_][\p{L}\p{N}_ ]*$/u.test(f.name)) errors.push(`Invalid field name "${f.name}" (letters, digits, spaces, _).`);
    if (names.has(f.name.toLowerCase())) errors.push(`Duplicate field "${f.name}".`);
    names.add(f.name.toLowerCase());
    if (!(FIELD_TYPES as readonly string[]).includes(f.type)) errors.push(`${f.name}: unknown type "${f.type}" (types: ${FIELD_TYPES.join(", ")}).`);
    if (f.type === "option" && !f.options?.length) errors.push(`${f.name}: a choice needs options.`);
    if (f.type === "option" && new Set(f.options?.map((o) => o.value)).size !== f.options?.length) errors.push(`${f.name}: duplicate options.`);
    if (f.type === "rating" && f.max !== undefined && ![5, 10].includes(f.max)) errors.push(`${f.name}: a rating goes up to 5 or 10.`);
    if (f.default !== undefined)
      try {
        coerce(f, f.default);
      } catch (e) {
        errors.push(`Default of ${e instanceof Error ? e.message : e}.`);
      }
    for (const [option, target] of Object.entries(f.branches ?? {})) {
      if (f.type !== "option" || f.multiple) errors.push(`${f.name}: only a single choice can branch.`);
      else if (!f.options?.some((o) => o.value === option)) errors.push(`${f.name}: branch on "${option}", which is not one of its options.`);
      if (target !== "end" && !sections.some((x) => x.title === target)) errors.push(`${f.name}: branch to "${target}", which is not a section (or "end").`);
    }
  }
  const titles = new Set<string>();
  for (const x of sections) {
    if (!x.title?.trim()) errors.push("A section needs a title.");
    if (titles.has(x.title)) errors.push(`Duplicate section "${x.title}".`);
    titles.add(x.title);
    for (const n of x.fields ?? []) {
      const f = fields.find((y) => y.name === n);
      if (!f) errors.push(`Section "${x.title}": no field "${n}".`);
      else if (f.form === false || isIdField(f)) errors.push(`Section "${x.title}": "${n}" is not asked in the form.`);
    }
    if (x.next && x.next !== "end" && !sections.some((y) => y.title === x.next)) errors.push(`Section "${x.title}": next "${x.next}" is not a section (or "end").`);
  }
  const placed = sections.flatMap((x) => x.fields ?? []);
  const twice = placed.filter((n, i) => placed.indexOf(n) !== i);
  if (twice.length) errors.push(`In several sections: ${[...new Set(twice)].join(", ")}.`);
  return errors;
}
