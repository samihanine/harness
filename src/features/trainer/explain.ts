/**
 * "Explain this visual": one plain request to the LLM (no tools) with what the visual shows — its
 * fields, its data as exported (first rows), the filters and slicer selections of its page, and for a
 * slicer its options and the model's description of its field — and the report's notes. Short
 * markdown in English, its form adapted to the visual and to what its data shows.
 */
import type { VisualDescriptor } from "powerbi-client";
import { AI_MODELS, openSession } from "@/agent/llm";
import { dataFacts } from "@/features/trainer/facts";
import { models, pageText, visualText, type Report } from "@/services/pbi";
import { settings, type DatasetEntry } from "@/lib/store";
import { EXPLAIN_PROMPT as PROMPT } from "@/prompts/trainer";

export async function explainVisual(report: Report, page: string, visual: string, context: string, dataset?: DatasetEntry) {
  const target = (await report.getPages()).find((p) => p.name === page);
  const v = (await target?.getVisuals())?.find((x) => x.name === visual);
  if (!v) throw new Error("Visual not found on the page.");
  const ref = dataset && { id: dataset.id, groupId: dataset.groupId };
  const m = await models();
  const [described, pageState, exported] = await Promise.all([visualText(v, true, false, ref), pageText(report, false), v.type === "slicer" ? undefined : v.exportData(m.ExportDataType.Summarized, 200).catch(() => undefined)]);
  const facts = exported?.data ? dataFacts(exported.data) : "";
  // A slicer has no data of its own: what its field is (model description, type, values) explains it.
  const fields = v.type === "slicer" ? await slicerFields(v, dataset) : "";
  const { model } = await settings.get();
  const prompt = [
    PROMPT,
    context && `<report>\n${context}\n</report>`,
    `<visual>\n${described}${fields ? `\n${fields}` : ""}\n</visual>`,
    facts && `<facts>\n${facts}\n</facts>`,
    `<page>\n${pageState}\n</page>`,
  ]
    .filter(Boolean)
    .join("\n\n");
  return asMarkdown(await openSession(model ?? AI_MODELS[0], "explain").send(prompt));
}

/**
 * The markdown of an answer: the inside of its ```markdown fence, HTML tags turned back into markdown
 * (some models answer in HTML whatever the prompt says).
 */
export function asMarkdown(text: string) {
  const fenced = text.match(/```(?:markdown|md)?\s*\n([\s\S]*?)```/i)?.[1] ?? text;
  return fenced
    .replace(/<\/?(b|strong)>/gi, "**")
    .replace(/<\/?(i|em)>/gi, "*")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<\/(li|p|div|h\d)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** The slicer's fields as the model describes them: table kind (e.g. a field parameter), type, description, values. */
async function slicerFields(v: VisualDescriptor, dataset?: DatasetEntry) {
  const state = await v.getSlicerState().catch(() => undefined);
  const lines: string[] = [];
  for (const t of (state?.targets ?? []) as { table?: string; column?: string; hierarchy?: string; hierarchyLevel?: string }[]) {
    const table = dataset?.info?.tables.find((x) => x.name === t.table);
    const name = t.column ?? t.hierarchyLevel;
    const column = table?.columns.find((c) => c.name === name);
    lines.push(
      [
        `  field '${t.table}'[${name}]${t.hierarchy ? ` (level of hierarchy ${t.hierarchy})` : ""}`,
        table?.kind && `table kind: ${table.kind}${table.extra ? ` (${table.extra})` : ""}`,
        column?.type && `type ${column.type}`,
        column?.description && `description: ${column.description}`,
        column?.values && `values: ${column.values}`,
      ]
        .filter(Boolean)
        .join("; "),
    );
  }
  return lines.join("\n");
}
