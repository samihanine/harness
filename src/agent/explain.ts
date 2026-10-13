/**
 * "Explain this visual": one plain request to the LLM (no tools) with what the visual shows — its
 * fields, its data as exported (first rows), the filters and slicer selections of its page — and the
 * report's notes. Short, factual, in the user's language.
 */
import { AI_MODELS, openSession } from "./llm";
import { pageText, visualText, type Report } from "@/lib/pbi";
import { settings } from "@/lib/store";

export async function explainVisual(report: Report, page: string, visual: string, context: string) {
  const target = (await report.getPages()).find((p) => p.name === page);
  const v = (await target?.getVisuals())?.find((x) => x.name === visual);
  if (!v) throw new Error("Visual not found on the page.");
  const [described, pageState] = await Promise.all([visualText(v, true), pageText(report, false)]);
  const { model } = await settings.get();
  const prompt = [
    "You explain one visual of a Power BI report to a business user. Use ONLY the data given below: never invent a figure. Reply in markdown, in the user's language (" + navigator.language + "), in 3 to 6 short bullet points: what the visual shows, the key figures, the trend or the notable gaps (highest, lowest, changes), and what the active filters mean for the reading. If the data is missing or truncated, say so briefly.",
    context && `<report>\n${context}\n</report>`,
    `<visual>\n${described}\n</visual>`,
    `<page>\n${pageState}\n</page>`,
  ]
    .filter(Boolean)
    .join("\n\n");
  return openSession(model ?? AI_MODELS[0], "explain").send(prompt);
}
