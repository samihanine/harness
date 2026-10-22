/** Pieces of the AI context shared by the screens. */
import { contextsText } from "@/lib/library";
import { modelSummary } from "@/lib/model";
import type { DatasetEntry, ReportEntry } from "@/lib/store";

/**
 * The context of a screen with the user's context files linked to its entries (company vocabulary,
 * rules…), given once at the start of a conversation (again only when it changes).
 */
export async function withContexts(parts: (string | false | undefined)[], contextIds: (string[] | undefined)[]) {
  const ids = [...new Set(contextIds.flatMap((x) => x ?? []))];
  const contexts = ids.length ? await contextsText(ids) : "";
  return [...parts, contexts && `Context files from the user (background knowledge; follow their rules):\n${contexts}`].filter(Boolean).join("\n\n");
}

/** What the AI knows of a report and its model: their names, the user's rights, the model's structure, the linked contexts. */
export const reportContext = (report: ReportEntry, dataset?: DatasetEntry) =>
  withContexts(
    [
      `Report "${report.name}" (id ${report.id}). The user's right on it: ${report.access ?? "unknown"}${report.accessNote ? ` (${report.accessNote})` : ""}.`,
      dataset
        ? `Semantic model "${dataset.name}" (right: ${dataset.access ?? "unknown"}${dataset.error ? `; ${dataset.error}` : ""})\n${dataset.info ? modelSummary(dataset.info) : "Structure not read."}`
        : "The semantic model is not known to the app: DAX and field checks are not available.",
    ],
    [report.contextIds, dataset?.contextIds],
  );
