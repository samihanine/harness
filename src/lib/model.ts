/**
 * What is known of a semantic model, read from the best source the user's rights allow:
 * 1. TMDL through Fabric (write access): everything — DAX, relationships, descriptions, formats,
 *    calculated columns, hierarchies, calculation groups, field parameters, roles;
 * 2. DAX INFO.VIEW functions (Build access): tables, columns, measures (often without DAX), relationships.
 * Then value samples per column (distinct text values, number / date ranges). Nothing blocks: each
 * failing part becomes a warning.
 */
import { dax, type Ref, type Rows } from "./pbi";
import { readModelFiles } from "./tmdl";

export type Column = { name: string; type?: string; description?: string; hidden?: boolean; format?: string; expression?: string; category?: string; sortBy?: string; values?: string };
export type Measure = { name: string; expression?: string; format?: string; description?: string; folder?: string; hidden?: boolean };
export type TableInfo = {
  name: string;
  description?: string;
  hidden?: boolean;
  kind?: "date" | "calculated" | "parameter" | "calculation group";
  rows?: number;
  columns: Column[];
  measures: Measure[];
  hierarchies: string[];
  /** Calculation group items / source query (short). */
  extra?: string;
};
export type Relationship = { from: string; to: string; cardinality: string; both: boolean; active: boolean };
export type ModelInfo = { source: "tmdl" | "info"; tables: TableInfo[]; relationships: Relationship[]; roles: string[]; warnings: string[]; readAt: string };

const AUTO_DATE = /^(LocalDateTable|DateTableTemplate)_/;
const message = (e: unknown) => (e instanceof Error ? e.message : String(e)).slice(0, 160);

export async function readModel(ref: Ref): Promise<ModelInfo> {
  const warnings: string[] = [];
  let info: ModelInfo;
  try {
    info = fromTmdl(await readModelFiles(ref.id, ref.groupId));
  } catch (e) {
    warnings.push(`Full definition not readable (${message(e)}): structure read with DAX, measure formulas may be missing.`);
    info = await fromInfoViews(ref, warnings);
  }
  info.warnings.unshift(...warnings);
  await addSamples(ref, info);
  return info;
}

/* ---------------------------------- TMDL ---------------------------------- */

type Node = { kind: string; name: string; expr?: string; props: Record<string, string>; children: Node[]; doc?: string };
const OBJECTS = new Set(["table", "column", "measure", "hierarchy", "level", "partition", "calculationGroup", "calculationItem", "relationship", "role", "variation", "annotation", "extendedProperty", "model", "database", "culture", "perspective", "expression", "dataSource", "tablePermission", "member", "linguisticMetadata", "formatStringDefinition", "detailRowsDefinition", "ref"]);
const unquote = (s: string) => (s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s);

/** Splits `'name with = sign' = expression` into name and expression. */
function header(rest: string) {
  let name = rest;
  let after = "";
  if (rest.startsWith("'")) {
    let i = 1;
    while (i < rest.length && !(rest[i] === "'" && rest[i + 1] !== "'")) i += rest[i] === "'" ? 2 : 1;
    name = rest.slice(0, i + 1);
    after = rest.slice(i + 1);
  } else {
    const eq = rest.indexOf(" =");
    if (eq >= 0) [name, after] = [rest.slice(0, eq), rest.slice(eq)];
  }
  const m = after.match(/^\s*=\s*(.*)$/);
  return { name: unquote(name.trim()), expr: m ? m[1] : undefined };
}

/** Indentation-based TMDL reader (objects, `key: value` properties, multi-line expressions, /// descriptions). */
function parseTmdl(text: string): Node[] {
  const roots: Node[] = [];
  const stack: { node: Node; depth: number }[] = [];
  let doc: string[] = [];
  let open: { target: { expr?: string } | Record<string, string>; key?: string; depth: number; fence?: boolean } | null = null;
  for (const raw of text.split(/\r?\n/)) {
    const depth = raw.match(/^\t*/)![0].length;
    const line = raw.trim();
    if (open) {
      const fenced = open.fence;
      if (fenced || (line === "" ? true : depth > open.depth)) {
        if (fenced && line === "```") {
          open = null;
          continue;
        }
        const piece = raw.slice(Math.min(depth, open.depth + 1));
        if (open.key) (open.target as Record<string, string>)[open.key] = `${(open.target as Record<string, string>)[open.key] ?? ""}${(open.target as Record<string, string>)[open.key] ? "\n" : ""}${piece}`;
        else (open.target as { expr?: string }).expr = `${(open.target as { expr?: string }).expr ?? ""}${(open.target as { expr?: string }).expr ? "\n" : ""}${piece}`;
        continue;
      }
      open = null;
    }
    if (!line) continue;
    if (line.startsWith("///")) {
      doc.push(line.slice(3).trim());
      continue;
    }
    const word = line.split(/\s/)[0];
    if (OBJECTS.has(word)) {
      while (stack.length && stack[stack.length - 1].depth >= depth) stack.pop();
      const { name, expr } = header(line.slice(word.length).trim());
      const node: Node = { kind: word, name, props: {}, children: [], doc: doc.join(" ") || undefined };
      doc = [];
      (stack.length ? stack[stack.length - 1].node.children : roots).push(node);
      stack.push({ node, depth });
      if (expr !== undefined) {
        node.expr = expr.startsWith("```") ? "" : expr.trim() || undefined;
        if (!expr.trim() || expr.startsWith("```")) open = { target: node, depth: depth + 1, fence: expr.startsWith("```") };
      }
      continue;
    }
    doc = [];
    const owner = [...stack].reverse().find((s) => s.depth < depth)?.node;
    if (!owner) continue;
    const prop = line.match(/^([\w]+)\s*(:|=)\s*(.*)$/);
    if (!prop) owner.props[line] = "true";
    else if (prop[2] === ":") owner.props[prop[1]] = prop[3];
    else {
      owner.props[prop[1]] = prop[3].startsWith("```") ? "" : prop[3];
      if (!prop[3].trim() || prop[3].startsWith("```")) open = { target: owner.props, key: prop[1], depth, fence: prop[3].startsWith("```") };
    }
  }
  return roots;
}

const short = (s?: string, n = 400) => (s && s.length > n ? `${s.slice(0, n)}…` : s);
const columnRef = (s: string) => {
  const m = s.match(/^('(?:[^']|'')+'|[^.]+)\.(.+)$/);
  return m ? { table: unquote(m[1]), column: unquote(m[2]) } : { table: "", column: s };
};

function fromTmdl(files: { path: string; text: string }[]): ModelInfo {
  const nodes = files.filter((f) => f.path.endsWith(".tmdl")).flatMap((f) => parseTmdl(f.text));
  const tables: TableInfo[] = nodes
    .filter((n) => n.kind === "table" && !AUTO_DATE.test(n.name))
    .map((t) => {
      const partition = t.children.find((c) => c.kind === "partition");
      const source = partition?.props.source ?? partition?.expr ?? "";
      const group = t.children.find((c) => c.kind === "calculationGroup");
      const kind: TableInfo["kind"] = group ? "calculation group" : /NAMEOF\s*\(/i.test(source) ? "parameter" : t.props.dataCategory === "Time" ? "date" : partition?.expr === "calculated" ? "calculated" : undefined;
      return {
        name: t.name,
        description: t.doc ?? t.props.description,
        hidden: t.props.isHidden === "true" || undefined,
        kind,
        columns: t.children
          .filter((c) => c.kind === "column")
          .map((c) => ({
            name: c.name,
            type: c.props.dataType,
            description: c.doc ?? c.props.description,
            hidden: c.props.isHidden === "true" || undefined,
            format: c.props.formatString,
            expression: short(c.expr),
            category: c.props.dataCategory,
            sortBy: c.props.sortByColumn,
          })),
        measures: t.children
          .filter((c) => c.kind === "measure")
          .map((m) => ({ name: m.name, expression: m.expr?.trim(), format: m.props.formatString, description: m.doc ?? m.props.description, folder: m.props.displayFolder, hidden: m.props.isHidden === "true" || undefined })),
        hierarchies: t.children.filter((c) => c.kind === "hierarchy").map((h) => `${h.name} (${h.children.filter((l) => l.kind === "level").map((l) => l.props.column ?? l.name).join(" › ")})`),
        extra: group
          ? `items: ${group.children.filter((c) => c.kind === "calculationItem").map((i) => `${i.name} = ${short(i.expr?.replace(/\s+/g, " "), 160)}`).join(" | ")}`
          : kind === "parameter" || kind === "calculated"
            ? `source: ${short(source.replace(/\s+/g, " "), 400)}`
            : undefined,
      };
    });
  const relationships = nodes
    .filter((n) => n.kind === "relationship" && n.props.fromColumn && n.props.toColumn)
    .map((r) => ({ from: columnRef(r.props.fromColumn), to: columnRef(r.props.toColumn), r }))
    .filter(({ from, to }) => !AUTO_DATE.test(from.table) && !AUTO_DATE.test(to.table))
    .map(({ from, to, r }) => ({
      from: `'${from.table}'[${from.column}]`,
      to: `'${to.table}'[${to.column}]`,
      cardinality: `${r.props.fromCardinality === "one" ? "1" : "*"}→${r.props.toCardinality === "many" ? "*" : "1"}`,
      both: r.props.crossFilteringBehavior === "bothDirections",
      active: r.props.isActive !== "false",
    }));
  const roles = nodes.filter((n) => n.kind === "role").map((r) => r.name);
  return { source: "tmdl", tables, relationships, roles, warnings: [], readAt: new Date().toISOString() };
}

/* -------------------------------- INFO.VIEW -------------------------------- */

async function fromInfoViews(ref: Ref, warnings: string[]): Promise<ModelInfo> {
  const read = (name: string) =>
    dax(ref, `EVALUATE ${name}()`).catch((e) => {
      warnings.push(`${name}: ${message(e)}`);
      return [] as Rows;
    });
  const [tables, columns, measures, relationships] = await Promise.all(["INFO.VIEW.TABLES", "INFO.VIEW.COLUMNS", "INFO.VIEW.MEASURES", "INFO.VIEW.RELATIONSHIPS"].map(read));
  // Nothing readable at all (no Build right): an error, not an empty model.
  if (!tables.length && !columns.length) throw new Error(`The model cannot be queried with DAX: ${warnings.at(-1) ?? "no table returned"}`);
  const byName = new Map<string, TableInfo>();
  const table = (name: string) => {
    if (!byName.has(name)) byName.set(name, { name, columns: [], measures: [], hierarchies: [] });
    return byName.get(name)!;
  };
  const text = (v: unknown) => (v === null || v === undefined || v === "" ? undefined : String(v));
  for (const t of tables) if (!AUTO_DATE.test(String(t.Name))) Object.assign(table(String(t.Name)), { description: text(t.Description), hidden: t.IsHidden === true || undefined, kind: t.DataCategory === "Time" ? "date" : undefined });
  for (const c of columns)
    if (!AUTO_DATE.test(String(c.Table)) && !String(c.Name).startsWith("RowNumber-"))
      table(String(c.Table)).columns.push({ name: String(c.Name), type: text(c.DataType), description: text(c.Description), hidden: c.IsHidden === true || undefined, format: text(c.FormatString), expression: short(text(c.Expression)), category: text(c.DataCategory) });
  for (const m of measures)
    table(String(m.Table)).measures.push({ name: String(m.Name), expression: text(m.Expression), format: text(m.FormatString), description: text(m.Description), folder: text(m.DisplayFolder), hidden: m.IsHidden === true || undefined });
  if (measures.length && measures.every((m) => !m.Expression)) warnings.push("Measure formulas are not readable with your rights (names only).");
  return {
    source: "info",
    tables: [...byName.values()],
    relationships: relationships
      .filter((r) => !AUTO_DATE.test(String(r.FromTable)) && !AUTO_DATE.test(String(r.ToTable)))
      .map((r) => ({
        from: `'${r.FromTable}'[${r.FromColumn}]`,
        to: `'${r.ToTable}'[${r.ToColumn}]`,
        cardinality: `${/one/i.test(String(r.FromCardinality)) ? "1" : "*"}→${/many/i.test(String(r.ToCardinality)) ? "*" : "1"}`,
        both: /both/i.test(String(r.CrossFilteringBehavior)),
        active: r.IsActive !== false,
      })),
    roles: [],
    warnings: [],
    readAt: new Date().toISOString(),
  };
}

/* --------------------------------- Samples --------------------------------- */

const SAMPLE_VALUES = 25;
const q = (s: string) => `'${s.replace(/'/g, "''")}'`;
const col = (t: string, c: string) => `${q(t)}[${c.replace(/]/g, "]]")}]`;
const isText = (type = "") => /^(string|text)$/i.test(type);
const isRange = (type = "") => /^(int64|integer|double|decimal|number|currency|datetime|date)$/i.test(type);

/** Row count, distinct text values (all when few, else examples), number / date ranges: one light query per table. */
async function addSamples(ref: Ref, info: ModelInfo) {
  const targets = info.tables.filter((t) => t.kind !== "calculation group" && t.kind !== "parameter" && t.columns.length).slice(0, 60);
  let failed = 0;
  const run = async (t: TableInfo) => {
    const cols = t.columns.filter((c) => !c.hidden && (isText(c.type) || isRange(c.type))).slice(0, 30);
    const parts = [`"rows", COUNTROWS(${q(t.name)})`];
    cols.forEach((c, i) => {
      const ref = col(t.name, c.name);
      if (isText(c.type)) parts.push(`"n${i}", COUNTROWS(DISTINCT(${ref}))`, `"v${i}", CONCATENATEX(TOPN(${SAMPLE_VALUES}, DISTINCT(${ref}), ${ref}, ASC), ${ref}, "|", ${ref}, ASC)`);
      else parts.push(`"a${i}", MIN(${ref})`, `"b${i}", MAX(${ref})`);
    });
    try {
      const [row] = await dax(ref, `EVALUATE ROW(${parts.join(", ")})`);
      t.rows = Number(row.rows) || 0;
      cols.forEach((c, i) => {
        if (isText(c.type)) {
          const n = Number(row[`n${i}`]) || 0;
          const values = String(row[`v${i}`] ?? "").split("|").filter(Boolean).map((v) => (v.length > 40 ? `${v.slice(0, 40)}…` : v));
          if (values.length) c.values = n <= SAMPLE_VALUES ? `${values.join(" | ")} (all ${n})` : `${n} distinct, e.g. ${values.slice(0, 6).join(" | ")}`;
        } else if (row[`a${i}`] !== null && row[`a${i}`] !== undefined) c.values = `${String(row[`a${i}`]).replace("T00:00:00", "")} … ${String(row[`b${i}`]).replace("T00:00:00", "")}`;
      });
    } catch {
      failed++;
    }
  };
  for (let i = 0; i < targets.length; i += 4) await Promise.all(targets.slice(i, i + 4).map(run));
  if (failed) info.warnings.push(`Value samples unavailable for ${failed} table(s).`);
}

/* ------------------------------ Views for the AI ------------------------------ */

const flags = (t: TableInfo) => [t.kind, t.hidden && "hidden", t.rows !== undefined && `${t.rows} rows`].filter(Boolean).join(", ");

/** Always in the AI context: tables (one line each), measure names by folder, relationships. Details through the model tools. */
export function modelSummary(info: ModelInfo) {
  const lines = [
    info.source === "tmdl" ? "Model read in full (DAX of every measure available)." : "Model read partially (with Build rights).",
    ...info.warnings.map((w) => `Limit: ${w}`),
    "Tables (details: describe_table; a measure: get_measure; search: search_model):",
  ];
  for (const t of info.tables) {
    const folders = new Map<string, string[]>();
    for (const m of t.measures.filter((m) => !m.hidden)) folders.set(m.folder ?? "", [...(folders.get(m.folder ?? "") ?? []), `[${m.name}]`]);
    const measures = [...folders].map(([f, names]) => (f ? `${f}: ${names.join(", ")}` : names.join(", "))).join("; ");
    lines.push(
      `- '${t.name}'${flags(t) ? ` (${flags(t)})` : ""}${t.description ? ` — ${short(t.description, 120)}` : ""}: ${t.columns.filter((c) => !c.hidden).length} columns${measures ? `; measures: ${measures}` : ""}`,
    );
  }
  if (info.relationships.length) lines.push("Relationships (many→one unless noted):", ...info.relationships.map((r) => `- ${r.from} ${r.cardinality} ${r.to}${r.both ? " (both directions)" : ""}${r.active ? "" : " (inactive: USERELATIONSHIP)"}`));
  if (info.roles.length) lines.push(`Security roles (RLS): ${info.roles.join(", ")}`);
  return lines.join("\n");
}

const measureLine = (m: Measure, table: string) =>
  [`[${m.name}] (in '${table}')${m.format ? ` format ${m.format}` : ""}${m.folder ? ` folder ${m.folder}` : ""}${m.hidden ? " hidden" : ""}`, m.description && `  description: ${m.description}`, `  = ${m.expression ?? "(formula not readable)"}`]
    .filter(Boolean)
    .join("\n");

export function describeTable(info: ModelInfo, name: string) {
  const t = info.tables.find((x) => x.name.toLowerCase() === name.replace(/^'|'$/g, "").toLowerCase());
  if (!t) return `No table "${name}". Tables: ${info.tables.map((x) => x.name).join(", ")}`;
  return [
    `'${t.name}'${flags(t) ? ` (${flags(t)})` : ""}${t.description ? `\n${t.description}` : ""}`,
    t.extra,
    "Columns:",
    ...t.columns.map(
      (c) =>
        `- ${c.name}: ${c.type ?? "?"}${c.format ? ` format ${c.format}` : ""}${c.category ? ` (${c.category})` : ""}${c.hidden ? " hidden" : ""}${c.sortBy ? ` sorted by ${c.sortBy}` : ""}${c.description ? ` — ${c.description}` : ""}${c.expression ? `\n    = ${c.expression}` : ""}${c.values ? `\n    values: ${c.values}` : ""}`,
    ),
    t.hierarchies.length ? `Hierarchies: ${t.hierarchies.join("; ")}` : "",
    t.measures.length ? "Measures:" : "",
    ...t.measures.map((m) => measureLine(m, t.name)),
    `Relationships: ${info.relationships.filter((r) => r.from.startsWith(`'${t.name}'`) || r.to.startsWith(`'${t.name}'`)).map((r) => `${r.from} ${r.cardinality} ${r.to}`).join("; ") || "none"}`,
  ]
    .filter(Boolean)
    .join("\n");
}

/** A measure with what it uses, what uses it, and the library reports showing it. */
export function describeMeasure(info: ModelInfo, name: string, reports: { name: string; snapshot?: string }[] = []) {
  const clean = name.replace(/^.*\[|\]$/g, "");
  const all = info.tables.flatMap((t) => t.measures.map((m) => ({ m, table: t.name })));
  const hit = all.find((x) => x.m.name.toLowerCase() === clean.toLowerCase());
  if (!hit) return `No measure "${name}". Use search_model to find it.`;
  const uses = all.filter((x) => x !== hit && hit.m.expression?.includes(`[${x.m.name}]`)).map((x) => `[${x.m.name}]`);
  const usedBy = all.filter((x) => x !== hit && x.m.expression?.includes(`[${hit.m.name}]`)).map((x) => `[${x.m.name}]`);
  const shownIn = reports.filter((r) => r.snapshot?.includes(`[${hit.m.name}]`)).map((r) => r.name);
  return [
    measureLine(hit.m, hit.table),
    uses.length && `uses measures: ${uses.join(", ")}`,
    usedBy.length && `used by measures: ${usedBy.join(", ")}`,
    shownIn.length && `shown in reports: ${shownIn.join(", ")}`,
  ]
    .filter(Boolean)
    .join("\n");
}

/** Tables, columns and measures whose name, description, values or DAX contain the query. */
export function searchModel(info: ModelInfo, query: string) {
  const qy = query.toLowerCase();
  const has = (...s: (string | undefined)[]) => s.some((x) => x?.toLowerCase().includes(qy));
  const out: string[] = [];
  for (const t of info.tables) {
    if (has(t.name, t.description)) out.push(`table '${t.name}'${t.description ? ` — ${short(t.description, 100)}` : ""}`);
    for (const c of t.columns) if (has(c.name, c.description, c.values, c.expression)) out.push(`column '${t.name}'[${c.name}] ${c.type ?? ""}${c.values ? ` values: ${short(c.values, 120)}` : ""}`);
    for (const m of t.measures) if (has(m.name, m.description, m.expression, m.folder)) out.push(`measure [${m.name}] in '${t.name}' = ${short(m.expression?.replace(/\s+/g, " "), 160) ?? "?"}`);
  }
  return out.length ? out.slice(0, 40).join("\n") + (out.length > 40 ? `\n… ${out.length - 40} more` : "") : `Nothing matches "${query}".`;
}

/** "Table[Measure]" names (to tell measures from columns when building visuals). */
export const measureNames = (info?: ModelInfo) => new Set(info?.tables.flatMap((t) => t.measures.map((m) => `${t.name}[${m.name}]`)) ?? []);

/** The whole structure as markdown: summary, then every table in detail (columns with sample values, measures with DAX, relationships). */
export function fullModelText(info: ModelInfo, name: string) {
  return [
    `# ${name}`,
    `Read ${info.readAt} (${info.source === "tmdl" ? "full definition" : "partial, with Build rights"})`,
    "",
    "## Summary",
    modelSummary(info),
    "",
    ...info.tables.flatMap((t) => [`## ${t.name}`, describeTable(info, t.name), ""]),
  ].join("\n");
}

/** Saves a text file in the browser. */
export function downloadText(name: string, text: string, type = "text/markdown") {
  const url = URL.createObjectURL(new Blob([text], { type }));
  Object.assign(document.createElement("a"), { href: url, download: name.replace(/[^\p{L}\p{N}.-]+/gu, "-") }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}
