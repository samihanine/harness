/** Pieces of the AI context shared by the screens. */
import { contextsText } from "@/lib/library";
import { modelSummary } from "@/lib/model";
import { modelSamples, profileText } from "@/lib/model-samples";
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

/**
 * A model for the AI: its structure (tables, columns, measures, relationships) and, when the user may
 * query it, examples of its data (values of each column, each measure's total).
 */
/** `focus`: the user's request (a big model then details the tables it is about). */
export async function datasetText(dataset: DatasetEntry, focus?: string) {
  const head = `Semantic model "${dataset.name}" (id ${dataset.id}; right: ${dataset.access ?? "unknown"}${dataset.error ? `; ${dataset.error}` : ""})`;
  if (!dataset.info) return `${head}\nStructure not read.`;
  // The value profile when the user made it (Datasets › "Value examples"), else a few random rows.
  const samples = dataset.info.profiledAt ? profileText(dataset.info, focus) : dataset.access === "view" ? "" : await modelSamples({ id: dataset.id, groupId: dataset.groupId }, dataset.info);
  return [head, modelSummary(dataset.info, focus), samples].filter(Boolean).join("\n");
}

/** What the AI knows of a report and its model: their names, the user's rights, the model's structure and data, the linked contexts. */
export const reportContext = async (report: ReportEntry, dataset?: DatasetEntry, focus?: string) =>
  withContexts(
    [
      `Report "${report.name}" (id ${report.id}). The user's right on it: ${report.access ?? "unknown"}${report.accessNote ? ` (${report.accessNote})` : ""}.`,
      dataset ? await datasetText(dataset, focus) : "The semantic model is not known to the app: DAX and field checks are not available.",
    ],
    [report.contextIds, dataset?.contextIds],
  );
