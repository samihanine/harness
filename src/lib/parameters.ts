/**
 * Parameters are tables of the semantic model (as Power BI Desktop creates them), so they need write
 * access to the model: written as TMDL, then the model is refreshed to compute them.
 */
import { readModelFiles, writeModelFiles } from "./tmdl";
import type { Ref } from "./pbi";

const q = (s: string) => `'${s.replace(/'/g, "''")}'`;
const str = (s: string) => `"${s.replace(/"/g, '""')}"`;

async function addTable(model: Ref, name: string, tmdl: string) {
  const files = await readModelFiles(model.id, model.groupId);
  if (files.some((f) => f.path === `definition/tables/${name}.tmdl`)) throw new Error(`A table "${name}" already exists.`);
  const next = files.map((f) => (f.path === "definition/model.tmdl" && /\nref table /.test(f.text) ? { ...f, text: `${f.text.trimEnd()}\nref table ${q(name)}\n` } : f));
  await writeModelFiles(model.id, [...next, { path: `definition/tables/${name}.tmdl`, text: tmdl }], model.groupId);
}

/** Field parameter: a slicer that switches the measures / columns shown by visuals. */
export async function addFieldParameter(model: Ref, name: string, fields: { label: string; field: string }[]) {
  const rows = fields.map((f, i) => {
    const m = f.field.match(/^'?([^'[]+)'?\[([^\]]+)\]$/);
    if (!m) throw new Error(`"${f.field}" must look like Table[Measure or Column]`);
    return `\t\t\t\t    (${str(f.label)}, NAMEOF(${q(m[1])}[${m[2]}]), ${i})`;
  });
  await addTable(
    model,
    name,
    `table ${q(name)}

\tcolumn ${q(name)}
\t\tdataType: string
\t\tsummarizeBy: none
\t\tsourceColumn: [Value1]
\t\tsortByColumn: ${q(`${name} Order`)}

\t\trelatedColumnDetails
\t\t\tgroupByColumn: ${q(`${name} Fields`)}

\tcolumn ${q(`${name} Fields`)}
\t\tdataType: string
\t\tisHidden
\t\tsummarizeBy: none
\t\tsourceColumn: [Value2]
\t\tsortByColumn: ${q(`${name} Order`)}

\t\textendedProperty ParameterMetadata =
\t\t\t\t{
\t\t\t\t  "version": 3,
\t\t\t\t  "kind": 2
\t\t\t\t}

\tcolumn ${q(`${name} Order`)}
\t\tdataType: int64
\t\tisHidden
\t\tformatString: 0
\t\tsummarizeBy: sum
\t\tsourceColumn: [Value3]

\tpartition ${q(name)} = calculated
\t\tmode: import
\t\tsource =
\t\t\t\t{
${rows.join(",\n")}
\t\t\t\t}
`,
  );
}

/** Numeric ("what-if") parameter: a slider and a measure holding the chosen value. */
export async function addNumericParameter(model: Ref, name: string, p: { min: number; max: number; step: number; value: number }) {
  await addTable(
    model,
    name,
    `table ${q(name)}

\tmeasure ${q(`${name} Value`)} = SELECTEDVALUE(${q(name)}[${name}], ${p.value})

\tcolumn ${q(name)}
\t\tdataType: double
\t\tsummarizeBy: none
\t\tsourceColumn: [Value]

\t\textendedProperty ParameterMetadata =
\t\t\t\t{
\t\t\t\t  "version": 0
\t\t\t\t}

\tpartition ${q(name)} = calculated
\t\tmode: import
\t\tsource = GENERATESERIES(${p.min}, ${p.max}, ${p.step})
`,
  );
}
