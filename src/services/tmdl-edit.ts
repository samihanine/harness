/**
 * Changes to a semantic model's definition (TMDL text): columns, measures, relationships, and
 * structured Power Query steps (each step is written by this code from a small spec, and the model's
 * columns follow it: a renamed, added or removed column changes in the table too).
 */
export type ModelFile = { path: string; text: string };

/* ------------------------------ Names and blocks ------------------------------ */

/** A TMDL name: quoted when it is not a plain word. */
export const q = (name: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
const unq = (s: string) => (s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s);
/** A DAX reference: 'Table'[Column]. */
export const daxRef = (table: string, column: string) => `'${table.replace(/'/g, "''")}'[${column.replace(/]/g, "]]")}]`;
/** An M text literal. */
const mText = (s: string) => `"${s.replace(/"/g, '""')}"`;

const OBJECT = /^\t(measure|column|partition|hierarchy|annotation|calculationGroup|changedProperty|extendedProperty)\s+(.*)$/;

type Block = { kind: string; name: string; lines: string[] };
type TableText = { head: string[]; blocks: Block[] };

/** Name of an object line ("'Name' = expr" or "Name"). */
function nameOf(rest: string) {
  if (rest.startsWith("'")) {
    let i = 1;
    while (i < rest.length && !(rest[i] === "'" && rest[i + 1] !== "'")) i += rest[i] === "'" ? 2 : 1;
    return unq(rest.slice(0, i + 1));
  }
  return rest.split(/\s*=|\s/)[0];
}

/** A table file as its header (table line and properties) and its objects (each with the /// lines before it). */
export function parseTable(text: string): TableText {
  const lines = text.replace(/\s+$/, "").split(/\r?\n/);
  const head: string[] = [];
  const blocks: Block[] = [];
  let doc: string[] = [];
  for (const line of lines) {
    const m = line.match(OBJECT);
    if (m) {
      blocks.push({ kind: m[1], name: nameOf(m[2]), lines: [...doc, line] });
      doc = [];
    } else if (/^\t\/\/\//.test(line)) doc.push(line);
    else if (blocks.length) {
      blocks[blocks.length - 1].lines.push(...doc, line);
      doc = [];
    } else head.push(...doc, line), (doc = []);
  }
  if (doc.length) (blocks.length ? blocks[blocks.length - 1].lines : head).push(...doc);
  return { head, blocks };
}

export const renderTable = (t: TableText) =>
  `${[t.head.join("\n").replace(/\s+$/, ""), ...t.blocks.map((b) => b.lines.join("\n").replace(/\s+$/, ""))].join("\n\n")}\n`;

const tableName = (t: TableText) => nameOf((t.head[0] ?? "").replace(/^table\s+/, ""));

/** An expression after "=": one line, or several lines indented under the object. */
const expressionLines = (expression: string) => {
  const lines = expression.trim().split(/\r?\n/);
  return lines.length === 1 ? [` ${lines[0]}`] : ["", ...lines.map((l) => `\t\t\t${l}`)];
};

/** Properties of a block (key: value lines at two tabs), set or removed (undefined). */
function setProps(block: Block, props: Record<string, string | undefined>) {
  for (const [key, value] of Object.entries(props)) {
    const i = block.lines.findIndex((l) => l.startsWith(`\t\t${key}:`) || l === `\t\t${key}`);
    if (value === undefined) {
      if (i >= 0) block.lines.splice(i, 1);
    } else if (i >= 0) block.lines[i] = `\t\t${key}: ${value}`;
    else {
      // After the expression lines (three tabs) and the other properties.
      let at = block.lines.findIndex((l, k) => k > 0 && /^\t\t\S/.test(l) && !/^\t\t(annotation|changedProperty) /.test(l));
      if (at < 0) at = block.lines.length;
      else while (at < block.lines.length && /^\t\t[a-zA-Z]+:/.test(block.lines[at])) at++;
      block.lines.splice(at, 0, `\t\t${key}: ${value}`);
    }
  }
}

/** The description (/// lines) of a block. */
function setDescription(block: Block, description?: string) {
  const start = block.lines.findIndex((l) => !/^\t\/\/\//.test(l));
  block.lines.splice(0, start, ...(description ? description.split("\n").map((l) => `\t/// ${l}`) : []));
}

/** The expression of a measure or calculated column (after "="). */
function setExpression(block: Block, keyword: string, name: string, expression: string) {
  const at = block.lines.findIndex((l) => OBJECT.test(l));
  let end = at + 1;
  while (end < block.lines.length && (/^\t\t\t/.test(block.lines[end]) || block.lines[end] === "")) end++;
  const [first, ...rest] = expressionLines(expression);
  block.lines.splice(at, end - at, `\t${keyword} ${q(name)} =${first}`, ...rest);
}

/* ---------------------------------- Model ---------------------------------- */

export class ModelDraft {
  files: Map<string, string>;
  constructor(files: ModelFile[]) {
    this.files = new Map(files.map((f) => [f.path, f.text]));
  }
  toFiles = (): ModelFile[] => [...this.files].map(([path, text]) => ({ path, text }));

  /** Path of a table's file, by its name (case ignored). */
  tablePath(name: string) {
    for (const [path, text] of this.files) if (path.startsWith("definition/tables/") && tableName(parseTable(text)).toLowerCase() === name.toLowerCase()) return path;
    throw new Error(`No table "${name}". Tables: ${this.tables().join(", ")}.`);
  }
  tables = () => [...this.files].filter(([p]) => p.startsWith("definition/tables/")).map(([, t]) => tableName(parseTable(t)));
  private edit(table: string, change: (t: TableText, name: string) => string) {
    const path = this.tablePath(table);
    const t = parseTable(this.files.get(path)!);
    const text = change(t, tableName(t));
    this.files.set(path, renderTable(t));
    return text;
  }
  private find(t: TableText, kind: string, name: string) {
    return t.blocks.find((b) => b.kind === kind && b.name.toLowerCase() === name.toLowerCase());
  }
  private need(t: TableText, kind: string, name: string) {
    const b = this.find(t, kind, name);
    if (!b) throw new Error(`No ${kind} "${name}" in '${tableName(t)}'. ${kind === "column" ? "Columns" : "Measures"}: ${t.blocks.filter((x) => x.kind === kind).map((x) => x.name).join(", ") || "none"}.`);
    return b;
  }

  /**
   * A field parameter: a calculated table listing fields (columns or measures, `ref` = 'Table'[Name])
   * that a slicer switches in the visuals using its column. Its values come with the next refresh.
   */
  addFieldParameter(name: string, items: { label: string; ref: string }[]) {
    if (this.tables().some((t) => t.toLowerCase() === name.toLowerCase())) throw new Error(`A table "${name}" already exists.`);
    if (items.length < 2) throw new Error("A field parameter lists at least 2 fields.");
    const fields = `${name} Fields`;
    const order = `${name} Order`;
    const rows = items.map((it, i) => `\t\t\t\t    ("${it.label.replace(/"/g, '""')}", NAMEOF(${it.ref}), ${i})`).join(",\n");
    const text = [
      `table ${q(name)}`,
      "",
      `\tcolumn ${q(name)}`,
      "\t\tdataType: string",
      "\t\tsummarizeBy: none",
      "\t\tsourceColumn: [Value1]",
      `\t\tsortByColumn: ${q(order)}`,
      "",
      "\t\trelatedColumnDetails",
      `\t\t\tgroupByColumn: ${q(fields)}`,
      "",
      `\tcolumn ${q(fields)}`,
      "\t\tdataType: string",
      "\t\tisHidden",
      "\t\tsummarizeBy: none",
      "\t\tsourceColumn: [Value2]",
      "",
      "\t\textendedProperty ParameterMetadata =",
      "\t\t\t\t{",
      '\t\t\t\t  "version": 3,',
      '\t\t\t\t  "kind": 2',
      "\t\t\t\t}",
      "",
      `\tcolumn ${q(order)}`,
      "\t\tdataType: int64",
      "\t\tisHidden",
      "\t\tsummarizeBy: sum",
      "\t\tsourceColumn: [Value3]",
      "",
      `\tpartition ${q(name)} = calculated`,
      "\t\tmode: import",
      "\t\tsource =",
      "\t\t\t\t{",
      rows,
      "\t\t\t\t}",
      "",
    ].join("\n");
    this.files.set(`definition/tables/${name.replace(/[\\/:*?"<>|]/g, "_")}.tmdl`, text);
    // The model lists its tables (their order) when it does: the new one is added there too.
    const model = this.files.get("definition/model.tmdl");
    if (model && /^ref table /m.test(model) && !new RegExp(`^ref table ${q(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`, "m").test(model)) this.files.set("definition/model.tmdl", `${model.replace(/\s+$/, "")}\nref table ${q(name)}\n`);
    return `Field parameter '${name}' created (${items.length} fields): put '${name}'[${name}] in a slicer, and in a visual's role instead of a field.`;
  }

  /* Measures */
  setMeasure(table: string, m: { name: string; expression?: string; format?: string; folder?: string; description?: string; hidden?: boolean }, create: boolean) {
    return this.edit(table, (t, tn) => {
      let b = this.find(t, "measure", m.name);
      if (create && b) throw new Error(`[${m.name}] already exists in '${tn}': use update_measure.`);
      if (!create && !b) b = this.need(t, "measure", m.name);
      if (!b) {
        if (!m.expression) throw new Error("A new measure needs its DAX expression.");
        b = { kind: "measure", name: m.name, lines: [] };
        // Measures first, like Power BI writes them.
        const first = t.blocks.findIndex((x) => x.kind !== "measure");
        t.blocks.splice(first < 0 ? t.blocks.length : first, 0, b);
        setExpression(b, "measure", m.name, m.expression);
      } else if (m.expression) setExpression(b, "measure", b.name, m.expression);
      setProps(b, { ...(m.format !== undefined && { formatString: m.format || undefined }), ...(m.folder !== undefined && { displayFolder: m.folder || undefined }), ...(m.hidden !== undefined && { isHidden: m.hidden ? "true" : undefined }) });
      if (m.description !== undefined) setDescription(b, m.description);
      return `Measure '${tn}'[${b.name}] ${create ? "created" : "changed"}.`;
    });
  }
  deleteObject(table: string, kind: "measure" | "column", name: string) {
    return this.edit(table, (t, tn) => {
      const b = this.need(t, kind, name);
      t.blocks.splice(t.blocks.indexOf(b), 1);
      return `${kind === "measure" ? "Measure" : "Column"} '${tn}'[${b.name}] deleted.`;
    });
  }

  /* Columns */
  /** A calculated (DAX) column. */
  addCalculatedColumn(table: string, name: string, expression: string, dataType?: string) {
    return this.edit(table, (t, tn) => {
      if (this.find(t, "column", name)) throw new Error(`'${tn}' already has a column "${name}".`);
      const b: Block = { kind: "column", name, lines: [] };
      setExpression(b, "column", name, expression);
      setProps(b, { ...(dataType && { dataType }), summarizeBy: "none" });
      const partition = t.blocks.findIndex((x) => x.kind === "partition");
      t.blocks.splice(partition < 0 ? t.blocks.length : partition, 0, b);
      return `Calculated column '${tn}'[${name}] added.`;
    });
  }
  /** A column read from the source (its query gives it). */
  addSourceColumn(t: TableText, name: string, dataType: string, sourceColumn = name) {
    // Already there (e.g. a step removed then added again): the same column, with its settings kept.
    const same = t.blocks.find((x) => x.kind === "column" && x.name.toLowerCase() === name.toLowerCase());
    if (same) return setProps(same, { dataType, sourceColumn });
    const b: Block = { kind: "column", name, lines: [`\tcolumn ${q(name)}`] };
    setProps(b, { dataType, summarizeBy: dataType === "double" || dataType === "int64" ? "sum" : "none", sourceColumn: sourceColumn, ...(dataType === "dateTime" && { formatString: "Short Date" }) });
    const partition = t.blocks.findIndex((x) => x.kind === "partition");
    t.blocks.splice(partition < 0 ? t.blocks.length : partition, 0, b);
  }
  updateColumn(table: string, name: string, p: { dataType?: string; format?: string; hidden?: boolean; description?: string; sortBy?: string; summarizeBy?: string; folder?: string; category?: string; expression?: string }) {
    return this.edit(table, (t, tn) => {
      const b = this.need(t, "column", name);
      if (p.expression !== undefined) {
        if (!/=\s*/.test(b.lines.find((l) => OBJECT.test(l)) ?? "") || b.lines.some((l) => l.startsWith("\t\tsourceColumn:"))) throw new Error(`'${tn}'[${b.name}] comes from the source: its values change with a Power Query step, not a DAX expression.`);
        setExpression(b, "column", b.name, p.expression);
      }
      if (p.sortBy && !this.find(t, "column", p.sortBy)) throw new Error(`No column "${p.sortBy}" in '${tn}' to sort by.`);
      setProps(b, {
        ...(p.dataType !== undefined && { dataType: p.dataType }),
        ...(p.format !== undefined && { formatString: p.format || undefined }),
        ...(p.hidden !== undefined && { isHidden: p.hidden ? "true" : undefined }),
        ...(p.sortBy !== undefined && { sortByColumn: p.sortBy ? q(p.sortBy) : undefined }),
        ...(p.summarizeBy !== undefined && { summarizeBy: p.summarizeBy }),
        ...(p.folder !== undefined && { displayFolder: p.folder || undefined }),
        ...(p.category !== undefined && { dataCategory: p.category || undefined }),
      });
      if (p.description !== undefined) setDescription(b, p.description);
      return `Column '${tn}'[${b.name}] changed.`;
    });
  }
  /** Renames a column in the model only (its source column keeps its name). */
  renameColumn(table: string, name: string, next: string) {
    return this.edit(table, (t, tn) => {
      const b = this.need(t, "column", name);
      if (this.find(t, "column", next)) throw new Error(`'${tn}' already has a column "${next}".`);
      const at = b.lines.findIndex((l) => OBJECT.test(l));
      b.lines[at] = b.lines[at].replace(/^\tcolumn\s+('(?:[^']|'')+'|[^\s=]+)/, `\tcolumn ${q(next)}`);
      b.name = next;
      return `Column '${tn}'[${name}] renamed to [${next}] (DAX using the old name must be updated).`;
    });
  }

  /* Relationships */
  private relationships() {
    const path = "definition/relationships.tmdl";
    const text = this.files.get(path) ?? "";
    const blocks = text.split(/\n(?=relationship )/).map((s) => s.trim()).filter((s) => s.startsWith("relationship "));
    return { path, blocks };
  }
  listRelationships() {
    return this.relationships().blocks.map((b) => {
      const prop = (k: string) => b.match(new RegExp(`\\n\\t${k}: (.*)`))?.[1]?.trim();
      return { name: b.split("\n")[0].replace(/^relationship\s+/, "").trim(), from: prop("fromColumn") ?? "", to: prop("toColumn") ?? "", both: prop("crossFilteringBehavior") === "bothDirections", active: prop("isActive") !== "false", fromCardinality: prop("fromCardinality") ?? "many", toCardinality: prop("toCardinality") ?? "one" };
    });
  }
  /** `from` is the many side ('Sales'[CustomerId]), `to` the one side ('Customer'[Id]). */
  addRelationship(r: { from: { table: string; column: string }; to: { table: string; column: string }; both?: boolean; active?: boolean; oneToOne?: boolean }) {
    const { path, blocks } = this.relationships();
    const ref = (c: { table: string; column: string }) => `${q(c.table)}.${q(c.column)}`;
    if (this.listRelationships().some((x) => (x.from === ref(r.from) && x.to === ref(r.to)) || (x.from === ref(r.to) && x.to === ref(r.from)))) throw new Error("These columns are already related.");
    const lines = [`relationship ${crypto.randomUUID()}`, ...(r.oneToOne ? ["\tfromCardinality: one"] : []), ...(r.both ? ["\tcrossFilteringBehavior: bothDirections"] : []), ...(r.active === false ? ["\tisActive: false"] : []), `\tfromColumn: ${ref(r.from)}`, `\ttoColumn: ${ref(r.to)}`];
    this.files.set(path, `${[...blocks, lines.join("\n")].join("\n\n")}\n`);
    return `Relationship ${ref(r.from)} (many) → ${ref(r.to)} (one) added${r.both ? ", both directions" : ""}${r.active === false ? ", inactive" : ""}.`;
  }
  deleteRelationship(from: string, to: string) {
    const { path, blocks } = this.relationships();
    const norm = (s: string) => s.replace(/['[\]\s]/g, "").replace(/\]/g, "").toLowerCase();
    const keep = blocks.filter((b) => {
      const f = b.match(/\n\tfromColumn: (.*)/)?.[1] ?? "", t = b.match(/\n\ttoColumn: (.*)/)?.[1] ?? "";
      return !((norm(f) === norm(from) && norm(t) === norm(to)) || (norm(f) === norm(to) && norm(t) === norm(from)));
    });
    if (keep.length === blocks.length) throw new Error(`No relationship between ${from} and ${to}.`);
    if (keep.length) this.files.set(path, `${keep.join("\n\n")}\n`);
    else this.files.delete(path);
    return "Relationship deleted.";
  }

  /* Power Query */
  /** The M query of a table (its partition's source) and its steps. */
  query(table: string) {
    const t = parseTable(this.files.get(this.tablePath(table))!);
    const partition = t.blocks.find((b) => b.kind === "partition");
    if (!partition || !/=\s*m\s*$/.test(partition.lines.find((l) => OBJECT.test(l)) ?? "")) throw new Error(`'${table}' has no Power Query (M) source.`);
    const at = partition.lines.findIndex((l) => /^\t\tsource\s*=/.test(l));
    let end = at + 1;
    while (end < partition.lines.length && (/^\t\t\t/.test(partition.lines[end]) || partition.lines[end] === "")) end++;
    const m = partition.lines.slice(at + 1, end).map((l) => l.replace(/^\t\t\t\t?/, "")).join("\n").trim();
    return { t, partition, at, end, m, steps: parseSteps(m) };
  }
  private setQuery(table: string, m: string) {
    return this.edit(table, (t) => {
      const partition = t.blocks.find((b) => b.kind === "partition")!;
      const at = partition.lines.findIndex((l) => /^\t\tsource\s*=/.test(l));
      let end = at + 1;
      while (end < partition.lines.length && (/^\t\t\t/.test(partition.lines[end]) || partition.lines[end] === "")) end++;
      partition.lines.splice(at + 1, end - at - 1, ...m.split("\n").map((l) => `\t\t\t\t${l}`));
      return "";
    });
  }
  /** Adds a step at the end of a table's query; the model's columns follow it. */
  addStep(table: string, step: Step) {
    const { steps } = this.query(table);
    if (!steps) throw new Error(`The query of '${table}' is not a let … in … expression: change it in Power BI Desktop.`);
    const previous = steps.result;
    let name = step.name?.replace(/[^\w ]/g, "").trim() || STEP_NAMES[step.kind];
    for (let i = 2; steps.list.some((s) => s.name === name); i++) name = `${STEP_NAMES[step.kind]} ${i}`;
    const id = /^[A-Za-z_]\w*$/.test(name) ? name : `#"${name}"`;
    const prevId = /^[A-Za-z_]\w*$/.test(previous) ? previous : `#"${previous}"`;
    const expression = stepM(step, prevId);
    const list = [...steps.list.map((s) => `    ${s.id} = ${s.expression}`), `    ${id} = ${expression}`];
    this.setQuery(table, `let\n${list.join(",\n")}\nin\n    ${id}`);
    // The columns of the model follow the step.
    const note = this.edit(table, (t, tn) => followStep(this, t, tn, step));
    return `Step "${name}" added to '${table}'.${note ? ` ${note}` : ""} The data changes after the next refresh.`;
  }
  /** Removes the LAST step of a table's query (the model's columns are not changed back). */
  removeLastStep(table: string) {
    const { steps } = this.query(table);
    if (!steps || steps.list.length < 2) throw new Error("No step to remove.");
    const list = steps.list.slice(0, -1);
    this.setQuery(table, `let\n${list.map((s) => `    ${s.id} = ${s.expression}`).join(",\n")}\nin\n    ${list.at(-1)!.id}`);
    return `Step "${steps.list.at(-1)!.name}" removed (check the columns: a column it added or renamed is still in the model).`;
  }
}

/* ------------------------------- Power Query steps ------------------------------- */

type MStep = { id: string; name: string; expression: string };
/** Steps of a "let … in …" query (top-level commas only). */
export function parseSteps(m: string): { list: MStep[]; result: string } | undefined {
  const match = m.match(/^let\s+([\s\S]*)\s+in\s+([\s\S]+)$/);
  if (!match) return undefined;
  const body = match[1];
  const parts: string[] = [];
  let depth = 0, inText = false, start = 0;
  for (let i = 0; i < body.length; i++) {
    const c = body[i];
    if (c === '"') {
      // "" inside a text is a quote, not its end.
      if (inText && body[i + 1] === '"') i++;
      else inText = !inText;
      continue;
    }
    if (inText) continue;
    if ("([{".includes(c)) depth++;
    else if (")]}".includes(c)) depth--;
    else if (c === "," && depth === 0) (parts.push(body.slice(start, i)), (start = i + 1));
  }
  parts.push(body.slice(start));
  const list = parts
    .map((p) => p.trim())
    .filter(Boolean)
    .map((p) => {
      const eq = p.match(/^(#"(?:[^"]|"")*"|[A-Za-z_][\w.]*)\s*=\s*([\s\S]*)$/);
      return eq ? { id: eq[1], name: eq[1].replace(/^#"(.*)"$/, "$1").replace(/""/g, '"'), expression: eq[2].trim() } : undefined;
    })
    .filter(Boolean) as MStep[];
  return { list, result: match[2].trim().replace(/^#"(.*)"$/, "$1") };
}

/** Structured steps (the AI fills them; this code writes the M). */
export type Step =
  | { kind: "rename_columns"; renames: { from: string; to: string }[]; name?: string }
  | { kind: "change_types"; types: { column: string; type: MType }[]; name?: string }
  | { kind: "filter_rows"; column: string; operator: "equals" | "not_equals" | "contains" | "not_contains" | "greater" | "less" | "not_empty" | "empty"; value?: string | number; name?: string }
  | { kind: "replace_values"; column: string; find: string; replace: string; whole?: boolean; name?: string }
  | { kind: "remove_columns"; columns: string[]; name?: string }
  | { kind: "conditional_column"; column: string; rules: { if_column: string; operator: "equals" | "contains" | "greater" | "less"; value: string | number; then: string | number }[]; otherwise: string | number | null; type?: MType; name?: string }
  | { kind: "custom_column"; column: string; m: string; type?: MType; name?: string }
  | { kind: "split_column"; column: string; delimiter: string; into: string[]; name?: string }
  | { kind: "transform_text"; columns: string[]; transform: "trim" | "upper" | "lower" | "proper" | "clean"; name?: string }
  | { kind: "fill_down"; columns: string[]; name?: string }
  | { kind: "remove_duplicates"; columns: string[]; name?: string }
  | { kind: "image_base64"; column: string; into: string; sources: ImageSource[]; name?: string };
/**
 * Where images are read. A site: its origin. SharePoint / OneDrive: the site (`path`, e.g.
 * /personal/<user> or /sites/<name>) and its libraries (`drives`: URL path → drive id), read through
 * the site's drives API — Power BI checks the credentials of a source on its address, and only the
 * API answers there with the user's token.
 */
export type ImageSource = { origin: string; path: string; sharepoint?: boolean; drives?: { path: string; id: string }[] };
export type MType = "text" | "number" | "integer" | "date" | "datetime" | "logical";
export const STEP_KINDS = ["rename_columns", "change_types", "filter_rows", "replace_values", "remove_columns", "conditional_column", "custom_column", "split_column", "transform_text", "fill_down", "remove_duplicates", "image_base64"] as const;
const STEP_NAMES: Record<Step["kind"], string> = {
  rename_columns: "Renamed Columns",
  change_types: "Changed Type",
  filter_rows: "Filtered Rows",
  replace_values: "Replaced Value",
  remove_columns: "Removed Columns",
  conditional_column: "Added Conditional Column",
  custom_column: "Added Custom",
  split_column: "Split Column",
  transform_text: "Transformed Text",
  fill_down: "Filled Down",
  remove_duplicates: "Removed Duplicates",
  image_base64: "Added Image Base64",
};
const M_TYPE: Record<MType, string> = { text: "type text", number: "type number", integer: "Int64.Type", date: "type date", datetime: "type datetime", logical: "type logical" };
const TMDL_TYPE: Record<MType, string> = { text: "string", number: "double", integer: "int64", date: "dateTime", datetime: "dateTime", logical: "boolean" };
const mValue = (v: string | number | null) => (v === null ? "null" : typeof v === "number" ? String(v) : mText(v));
const col = (c: string) => `[${c.replace(/]/g, "]]")}]`;
const list = (cols: string[]) => `{${cols.map(mText).join(", ")}}`;
const condition = (column: string, operator: string, value?: string | number) => {
  const c = col(column);
  switch (operator) {
    case "equals":
      return `${c} = ${mValue(value ?? null)}`;
    case "not_equals":
      return `${c} <> ${mValue(value ?? null)}`;
    case "contains":
      return `Text.Contains(${c}, ${mValue(String(value ?? ""))})`;
    case "not_contains":
      return `not Text.Contains(${c}, ${mValue(String(value ?? ""))})`;
    case "greater":
      return `${c} > ${mValue(value ?? 0)}`;
    case "less":
      return `${c} < ${mValue(value ?? 0)}`;
    case "empty":
      return `${c} = null or ${c} = ""`;
    default:
      return `${c} <> null and ${c} <> ""`;
  }
};

function stepM(s: Step, prev: string) {
  switch (s.kind) {
    case "rename_columns":
      return `Table.RenameColumns(${prev}, {${s.renames.map((r) => `{${mText(r.from)}, ${mText(r.to)}}`).join(", ")}})`;
    case "change_types":
      return `Table.TransformColumnTypes(${prev}, {${s.types.map((t) => `{${mText(t.column)}, ${M_TYPE[t.type]}}`).join(", ")}})`;
    case "filter_rows":
      return `Table.SelectRows(${prev}, each ${condition(s.column, s.operator, s.value)})`;
    case "replace_values":
      return `Table.ReplaceValue(${prev}, ${mText(s.find)}, ${mText(s.replace)}, ${s.whole ? "Replacer.ReplaceValue" : "Replacer.ReplaceText"}, {${mText(s.column)}})`;
    case "remove_columns":
      return `Table.RemoveColumns(${prev}, ${list(s.columns)})`;
    case "conditional_column": {
      const body = s.rules.reduceRight((rest, r) => `if ${condition(r.if_column, r.operator, r.value)} then ${mValue(r.then)} else ${rest}`, mValue(s.otherwise));
      return `Table.AddColumn(${prev}, ${mText(s.column)}, each ${body}, ${M_TYPE[s.type ?? "text"]})`;
    }
    case "custom_column":
      if (dynamicSource(s.m))
        throw new Error("This expression reads a source whose address comes from a column: Power BI cannot refresh such a query (DynamicDataSourcesIsNotSupportedForRefresh). To read images from a URL column, use step_image_base64; for other files, the host must be written as text: Web.Contents(\"https://host\", [RelativePath = …]).");
      return `Table.AddColumn(${prev}, ${mText(s.column)}, each ${s.m}, ${M_TYPE[s.type ?? "text"]})`;
    case "split_column":
      return `Table.SplitColumn(${prev}, ${mText(s.column)}, Splitter.SplitTextByDelimiter(${mText(s.delimiter)}, QuoteStyle.Csv), ${list(s.into)})`;
    case "transform_text": {
      const fn = { trim: "Text.Trim", upper: "Text.Upper", lower: "Text.Lower", proper: "Text.Proper", clean: "Text.Clean" }[s.transform];
      return `Table.TransformColumns(${prev}, {${s.columns.map((c) => `{${mText(c)}, ${fn}, type text}`).join(", ")}})`;
    }
    case "fill_down":
      return `Table.FillDown(${prev}, ${list(s.columns)})`;
    case "remove_duplicates":
      return `Table.Distinct(${prev}, ${list(s.columns)})`;
    case "image_base64":
      return imageBase64(s, prev);
  }
}

/** Power BI keeps 32,766 characters per text value: "data:image/jpeg;base64," + the base64 of at most this many bytes. */
export const IMAGE_MAX_BYTES = 24_400;
/** SharePoint thumbnail sizes tried in turn (bounding boxes, JPEG): the first that fits is kept. */
const THUMBNAILS = ["c400x400", "c320x320", "c240x240", "c160x160"];

/**
 * Images read from their URL, as "data:image/…;base64,…" (what Power BI shows in tables and cards).
 * SharePoint / OneDrive files (any size, any site or library) are read as a thumbnail made by
 * SharePoint (`/_api/v2.0/shares/<url>/driveItem/thumbnails`): a 3 MB photo comes as ~20 KB, small
 * enough for Power BI's text limit and quick to refresh; the largest size that fits is kept. Other
 * sites are read as they are (an image above the limit gives no image). Power BI refreshes only
 * sources whose host is written in the query: one Web.Contents per host, the rest as RelativePath /
 * Query. A URL of another host, or a file that cannot be read, gives no image (never a failed refresh).
 */
function imageBase64(s: Extract<Step, { kind: "image_base64" }>, prev: string) {
  if (!s.sources.length) throw new Error(`No image URL in [${s.column}] yet: nothing to read.`);
  const sources = [...s.sources].sort((a, b) => b.path.length - a.path.length);
  const read = sources.reduceRight((rest, x) => {
    if (!x.sharepoint) return `if h = ${mText(x.origin)} then Binary.Buffer(Web.Contents(${mText(x.origin)}, [RelativePath = Text.TrimStart(p[Path], "/"), Query = p[Query]])) else ${rest}`;
    // The site's drives API: it answers the user's token (Power BI tests the credentials on this address).
    const api = `${x.origin}${encodeURI(x.path)}/_api/v2.0/drives`;
    const libs = [...(x.drives ?? [])].sort((a, b) => b.path.length - a.path.length).map((d) => `{${mText(d.path)}, ${mText(d.id)}}`);
    return `if h = ${mText(x.origin)} and Text.StartsWith(p[Path], ${mText(`${x.path}/`)}) then let lib = List.First(List.Select({${libs.join(", ")}}, (l) => Text.StartsWith(p[Path], l{0} & "/"))), file = if lib = null then null else Text.Combine(List.Transform(Text.Split(Text.Range(p[Path], Text.Length(lib{0}) + 1), "/"), Uri.EscapeDataString), "/") in if lib = null then null else fit((size) => Web.Contents(${mText(api)}, [RelativePath = lib{1} & "/root:/" & file & ":/thumbnails/0/" & size & "/content"]), {${THUMBNAILS.map(mText).join(", ")}}) else ${rest}`;
  }, "null");
  const fit = `fit = (get, sizes) => let b = Binary.Buffer(get(List.First(sizes))) in if Binary.Length(b) <= ${IMAGE_MAX_BYTES} or List.Count(sizes) = 1 then b else @fit(get, List.Skip(sizes, 1))`;
  // The type from the file's first bytes (thumbnails are JPEG whatever the source), else its extension.
  const kind = `b4 = Binary.ToList(Binary.Range(bin, 0, 4)), head = Text.Lower(Text.FromBinary(Binary.Range(bin, 0, 64))), kind = if List.Count(b4) < 4 then ext else if b4{0} = 255 and b4{1} = 216 then "jpeg" else if b4{0} = 137 and b4{1} = 80 then "png" else if b4{0} = 71 and b4{1} = 73 then "gif" else if b4{0} = 82 and b4{1} = 73 then "webp" else if Text.Contains(head, "<svg") or Text.StartsWith(head, "<?xml") then "svg+xml" else ext`;
  return `Table.AddColumn(${prev}, ${mText(s.into)}, each let u = ${col(s.column)}, p = if u = null or u = "" then null else Uri.Parts(u), h = if p = null then null else p[Scheme] & "://" & p[Host], ${fit}, e = if p = null then "" else Text.Lower(List.Last(Text.Split(p[Path], "."))), ext = if e = "jpg" then "jpeg" else if e = "svg" then "svg+xml" else if List.Contains({"png", "jpeg", "gif", "webp"}, e) then e else "png", bin = if p = null then null else try ${read} otherwise null, ${kind} in if bin = null or Binary.Length(bin) > ${IMAGE_MAX_BYTES} then null else "data:image/" & kind & ";base64," & Binary.ToText(bin, BinaryEncoding.Base64), type text)`;
}

/**
 * The places to read images from, for these URLs: per site its origin; on SharePoint / OneDrive, per
 * site (/personal/x, /sites/x, /teams/x, or the root site), whose libraries (`drives`) are added by the
 * caller (they need an API call).
 */
export function imageSources(urls: string[]): ImageSource[] {
  const out = new Map<string, ImageSource>();
  for (const u of urls) {
    try {
      const url = new URL(u);
      if (!/\.sharepoint\.com$/i.test(url.host)) {
        out.set(url.origin, { origin: url.origin, path: "" });
        continue;
      }
      const parts = decodeURIComponent(url.pathname).split("/").filter(Boolean);
      const path = /^(personal|sites|teams)$/i.test(parts[0] ?? "") && parts[1] ? `/${parts[0]}/${parts[1]}` : "";
      out.set(url.origin + path, { origin: url.origin, path, sharepoint: true });
    } catch {
      // not a URL: no image
    }
  }
  return [...out.values()];
}

/** An M expression Power BI could not refresh: Web.Contents / other sources on an address built from a column. */
export function dynamicSource(m: string) {
  return /\b(Web\.Contents|Web\.Page|SharePoint\.Files|SharePoint\.Contents|OData\.Feed)\s*\(\s*(?!")/.test(m);
}

/** The table's columns after a step: renamed, typed, added or removed like the query's. */
function followStep(draft: ModelDraft, t: TableText, tn: string, s: Step) {
  const byColumn = (name: string) => t.blocks.find((b) => b.kind === "column" && (b.lines.find((l) => l.startsWith("\t\tsourceColumn:"))?.split(": ")[1]?.trim() ?? b.name) === name);
  const setSource = (b: Block, source: string) => {
    const i = b.lines.findIndex((l) => l.startsWith("\t\tsourceColumn:"));
    if (i >= 0) b.lines[i] = `\t\tsourceColumn: ${source}`;
  };
  switch (s.kind) {
    case "rename_columns": {
      for (const r of s.renames) {
        const b = byColumn(r.from);
        if (!b) continue;
        setSource(b, r.to);
        if (b.name === r.from) {
          const at = b.lines.findIndex((l) => OBJECT.test(l));
          b.lines[at] = `\tcolumn ${q(r.to)}`;
          b.name = r.to;
        }
      }
      return "Model columns renamed with it (DAX using the old names must be updated).";
    }
    case "change_types":
      for (const x of s.types) {
        const b = byColumn(x.column);
        if (b) setProps(b, { dataType: TMDL_TYPE[x.type], ...(x.type === "date" && { formatString: "Short Date" }) });
      }
      return "";
    case "remove_columns":
      for (const c of s.columns) {
        const b = byColumn(c);
        if (b) t.blocks.splice(t.blocks.indexOf(b), 1);
      }
      return "Their model columns are removed too.";
    case "conditional_column":
    case "custom_column":
      draft.addSourceColumn(t, s.column, TMDL_TYPE[s.type ?? "text"]);
      return `Column [${s.column}] added to the model.`;
    case "image_base64": {
      draft.addSourceColumn(t, s.into, "string");
      const b = t.blocks.find((x) => x.kind === "column" && x.name === s.into);
      if (b) setProps(b, { dataCategory: "ImageUrl" });
      const sp = s.sources.filter((x) => x.sharepoint);
      const libraries = sp.flatMap((x) => x.drives ?? []);
      const other = s.sources.filter((x) => !x.sharepoint);
      return `Column [${s.into}] added to the model (image). ${sp.length ? `SharePoint / OneDrive images (libraries: ${libraries.map((d) => d.path).join(", ") || "none found"}) are read as thumbnails of at most 400×400 px (largest that fits Power BI's text limit): any size of file works. ` : ""}${other.length ? `Images of ${other.map((x) => x.origin).join(", ")} are read as they are: above ${Math.round(IMAGE_MAX_BYTES / 1000)} KB they give no image (put them on SharePoint to have them resized). ` : ""}`.trim();
    }
    case "split_column": {
      const b = byColumn(s.column);
      if (b) t.blocks.splice(t.blocks.indexOf(b), 1);
      for (const c of s.into) draft.addSourceColumn(t, c, "string");
      return `Columns ${s.into.map((c) => `[${c}]`).join(", ")} replace [${s.column}] in the model.`;
    }
    default:
      void tn;
      return "";
  }
}
