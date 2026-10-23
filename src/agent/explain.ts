/**
 * "Explain this visual": one plain request to the LLM (no tools) with what the visual shows — its
 * fields, its data as exported (first rows), the filters and slicer selections of its page, and for a
 * slicer its options and the model's description of its field — and the report's notes. Short
 * markdown in English, its form adapted to the visual and to what its data shows.
 */
import type { VisualDescriptor } from "powerbi-client";
import { AI_MODELS, openSession } from "./llm";
import { pageText, visualText, type Report } from "@/lib/pbi";
import { settings, type DatasetEntry } from "@/lib/store";

const PROMPT = `You are a sharp data analyst. A business user clicked ONE visual of a Power BI report and asks: "what does this tell me?". Below are the visual (type, fields, filters, its data as exported), the page it sits on (other visuals, filters, slicers) and notes about the report.

Your job is to make the visual SPEAK: tell the user what they would only see after studying it for a few minutes.
- Find what matters in the data: the leader and how far ahead it is, the concentration (e.g. "the top 3 make 72 %"), the gap between best and worst, the trend and where it breaks, outliers, what is unexpectedly low or missing, a value that changes the story once compared with the total or another visual of the page. Compute these yourself (shares, differences, ratios, averages) from the rows given.
- Fit the form to the visual and the data — no fixed template:
  · a time series: trend, peak, dip, turning point, the latest value vs the usual level;
  · a ranking (bar chart, table): leader, tail, concentration, an item that stands out;
  · a card / KPI: put the number in context (share of what, compared to what on the page);
  · a matrix: the cell or row that breaks the pattern;
  · a slicer: what it filters, what is selected now and what that hides, which choice is worth trying;
  · a pie / donut: the dominant part and whether the split is balanced.
- Say only what this data supports. Do not restate the filters, the slicer selections or the page: the user sees them; mention a filter only if it changes the meaning.
- Lead with the single most useful insight in one bold sentence. Then at most 3 short points, only if they add something. Short: 25 to 70 words.
- Round numbers (661.8M, 12.5K, +8 %, 3×). Never list every value, never restate the title or the axes, no generic advice ("monitor this metric"), no hedging filler.
- Answer in the language of the report's titles (French titles → French), as MARKDOWN inside a \`\`\`markdown fence: the bold sentence on its own line, then one point per line starting with "- ", nothing else (no title, no closing sentence).

Good:
\`\`\`markdown
**Two regions make 81 % of the revenue: East (1.1B) leads West by 2×.**
- North is last (4 %), half of last year.
\`\`\`
Bad: "This chart shows revenue by region. East has the highest revenue."`;

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
