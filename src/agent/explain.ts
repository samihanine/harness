/**
 * "Explain this visual": one plain request to the LLM (no tools) with what the visual shows — its
 * fields, its data as exported (first rows), the filters and slicer selections of its page, and for a
 * slicer its options and the model's description of its field — and the report's notes. Short,
 * structured markdown, in English (the examples in the prompt set the length and the form).
 */
import type { VisualDescriptor } from "powerbi-client";
import { AI_MODELS, openSession } from "./llm";
import { pageText, visualText, type Report } from "@/lib/pbi";
import { settings, type DatasetEntry } from "@/lib/store";

const PROMPT = `You explain ONE visual of a Power BI report to a business user, from the data given below only (never invent a figure).
Reply in English, in markdown, SHORT: at most 5 bullet points and about 80 words. Round numbers (661.8M, 12.5K, +8%). Name only the figures that matter (highest, lowest, biggest change); never list every value. Mention filters only when they change the reading.

Example for a chart:
**Monthly sales, this year vs last year.**
- **Peak:** May, 661.8M (+8% vs LY).
- **Low:** October, 214.7M.
- **Trend:** above last year in Jan–Feb and May–Jul, below in March (−7%).
- **Scope:** current year only; Central and Unknown zones excluded.

Example for a slicer:
**Filters the page by season.**
- **Options:** Fall, Winter, Summer, Cruise, Permanent, Others.
- **Now:** nothing selected, so every season is shown.
- **Tip:** pick one season to compare its sales with the rest.

Example for a card:
**Total revenue of the selection: 1.97B.**
- Covers all products and regions (no filter).`;

export async function explainVisual(report: Report, page: string, visual: string, context: string, dataset?: DatasetEntry) {
  const target = (await report.getPages()).find((p) => p.name === page);
  const v = (await target?.getVisuals())?.find((x) => x.name === visual);
  if (!v) throw new Error("Visual not found on the page.");
  const ref = dataset && { id: dataset.id, groupId: dataset.groupId };
  const [described, pageState] = await Promise.all([visualText(v, true, false, ref), pageText(report, false)]);
  // A slicer has no data of its own: what its field is (model description, type, values) explains it.
  const fields = v.type === "slicer" ? await slicerFields(v, dataset) : "";
  const { model } = await settings.get();
  const prompt = [
    PROMPT,
    context && `<report>\n${context}\n</report>`,
    `<visual>\n${described}${fields ? `\n${fields}` : ""}\n</visual>`,
    `<page>\n${pageState}\n</page>`,
  ]
    .filter(Boolean)
    .join("\n\n");
  return openSession(model ?? AI_MODELS[0], "explain").send(prompt);
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
