/** Measures written into a model's TMDL (the table's file), the rest of the file untouched. */
import type { ModelFile } from "./definitions";

const quote = (name: string) => (/^[A-Za-z_][\w]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
const unquote = (s: string) => (s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s);

/** [start, end) lines of a measure block in a table file (end: the next object at the same depth). */
function measureRange(lines: string[], name: string) {
  const start = lines.findIndex((l) => {
    const m = l.match(/^\tmeasure\s+('(?:[^']|'')+'|[^\s=]+)/);
    return m && unquote(m[1]).toLowerCase() === name.toLowerCase();
  });
  if (start < 0) return undefined;
  let end = start + 1;
  while (end < lines.length && (lines[end].trim() === "" || /^\t\t/.test(lines[end]))) end++;
  // Blank lines at the end belong to the next object.
  while (end > start + 1 && lines[end - 1].trim() === "") end--;
  return { start, end };
}

export function tableFile(files: ModelFile[], table: string) {
  const f = files.find((x) => x.path.endsWith(".tmdl") && new RegExp(`^table\\s+('${table.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/'/g, "''")}'|${table.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")})\\s*$`, "m").test(x.text));
  if (!f) throw new Error(`No table "${table}" in the model's definition.`);
  return f;
}

export function upsertMeasure(file: ModelFile, m: { name: string; expression: string; format?: string; folder?: string; description?: string }) {
  const lines = file.text.replace(/\r/g, "").split("\n");
  const expr = m.expression.trim().split("\n");
  const block = [
    ...(m.description ? m.description.split("\n").map((d) => `\t/// ${d}`) : []),
    expr.length === 1 ? `\tmeasure ${quote(m.name)} = ${expr[0]}` : `\tmeasure ${quote(m.name)} =`,
    ...(expr.length > 1 ? expr.map((l) => `\t\t\t${l}`) : []),
    ...(m.format ? [`\t\tformatString: ${m.format}`] : []),
    ...(m.folder ? [`\t\tdisplayFolder: ${m.folder}`] : []),
  ];
  const range = measureRange(lines, m.name);
  if (range) {
    // Descriptions above the old block are replaced too.
    let from = range.start;
    while (from > 0 && /^\t\/\/\//.test(lines[from - 1])) from--;
    // Properties the call does not give (lineageTag, annotations…) are kept.
    const kept = lines.slice(range.start + 1, range.end).filter((l) => /^\t\t\w+\s*:/.test(l) && !/^\t\t(formatString|displayFolder)\s*:/.test(l) && !(m.format && /formatString/.test(l)));
    lines.splice(from, range.end - from, ...block, ...kept);
  } else {
    // A new measure: after the table line and its properties, before the first column.
    const at = lines.findIndex((l, i) => i > 0 && /^\t(column|measure|hierarchy|partition)\b/.test(l));
    lines.splice(at < 0 ? lines.length : at, 0, ...block, "");
  }
  return { ...file, text: lines.join("\n") };
}

export function deleteMeasure(file: ModelFile, name: string) {
  const lines = file.text.replace(/\r/g, "").split("\n");
  const range = measureRange(lines, name);
  if (!range) throw new Error(`No measure "${name}" in this table.`);
  let from = range.start;
  while (from > 0 && /^\t\/\/\//.test(lines[from - 1])) from--;
  lines.splice(from, range.end - from);
  return { ...file, text: lines.join("\n") };
}
