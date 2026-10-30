/**
 * The Builder's agent: it changes a report's definition (PBIR) without limit — pages, visuals,
 * formatting, theme, report measures — in a draft saved when its answer is done.
 */
import { reportContext } from "@/agent/context";
import { examplesText, referenceReports } from "@/services/report-schema";
import { PBIR_GUIDE, draftProblems, editorState, editorTools } from "./editor-tools";
import { draftInfo, fieldParameterTool, saveModelDraft } from "@/features/datasets/model-editor-tools";
import type { Agent } from "@/agent/loop";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES } from "@/agent/pbi-tools";
import type { Report } from "@/services/pbi";
import { pendingFiles, saveDraft } from "@/features/builder/report-draft";
import type { DatasetEntry, ReportEntry } from "@/lib/store";
import { BUILDER_INSTRUCTIONS } from "@/prompts/builder";

/** The visuals a draft writes, by page (from its files: definition/pages/<page>/visuals/<visual>/visual.json). */
const writtenVisuals = (files: string[]) =>
  files.flatMap((f) => {
    const m = f.match(/pages\/([^/]+)\/visuals\/([^/]+)\/visual\.json/);
    return m ? [{ page: m[1], visual: m[2] }] : [];
  });

/**
 * The editor reloaded until it shows what was written: the service serves the new definition a little
 * after it is saved, and an editor reloaded too early kept the old report — the new visuals did not
 * show, and a save from the editor then wrote the old report back over them.
 */
async function showSaved(editor: Report | undefined, written: { page: string; visual: string }[]) {
  if (!editor) return;
  for (let attempt = 0; attempt < 5; attempt++) {
    await new Promise((r) => setTimeout(r, 1500 + attempt * 1500));
    await editor.reload().catch(() => undefined);
    await new Promise((r) => setTimeout(r, 1500));
    if (!written.length) return;
    const pages = await editor.getPages().catch(() => []);
    const shown = new Set<string>();
    for (const page of pages.filter((p) => written.some((w) => w.page === p.name))) for (const v of await page.getVisuals().catch(() => [])) shown.add(`${page.name}/${v.name}`);
    if (written.every((w) => shown.has(`${w.page}/${w.visual}`) || !pages.some((p) => p.name === w.page))) return;
  }
}

/** `report`: the embedded editor when there is one (saved before the definition is written, reloaded after). */
export function builderAgent({ entry, dataset, report, onSaved }: { entry: () => ReportEntry | undefined; dataset: () => DatasetEntry | undefined; report: () => Report | undefined; onSaved?: () => void }): Agent {
  const ref = () => entry() && { id: entry()!.id, groupId: entry()!.groupId };
  const save = async () => report()?.save();
  const datasetRef = () => entry() && { id: entry()!.datasetId, groupId: dataset()?.groupId ?? entry()!.datasetGroupId };
  const ctx = {
    report,
    ref,
    dataset: datasetRef,
    // With the model's unsaved changes (a field parameter made in this answer is usable at once).
    info: () => (datasetRef() && draftInfo(datasetRef()!)) ?? dataset()?.info,
    references: () => referenceReports(entry()?.exampleIds),
  };
  /** The draft is checked once before it is saved: what is wrong goes back to the AI (saved anyway the second time). */
  let checked = false;
  return {
    instructions: BUILDER_INSTRUCTIONS,
    tools: [...editorTools(ctx, save), ...modelTools(dataset), ...(dataset()?.access === "edit" ? [fieldParameterTool(datasetRef)] : [])],
    // The reference reports the user chose (Reports tab) come with the context.
    context: async (focus) => [await reportContext(entry()!, dataset(), focus), await examplesText(entry()?.exampleIds)].filter(Boolean).join("\n\n"),
    state: () => editorState(ctx, save),
    finish: async () => {
      const r = ref();
      const d = datasetRef();
      // The model first: the report's visuals may use what was added to it (a field parameter).
      const model = d ? await saveModelDraft(d) : undefined;
      if (!r || !pendingFiles(r).length) return model && `✓ ${model}`;
      if (!checked) {
        checked = true;
        const problems = await draftProblems(ctx, save);
        if (problems.length) throw new Error(`The check of your draft found problems (fix them, e.g. layout_page for the layout, update_measure for a measure; then answer):\n${problems.map((x) => `- ${x}`).join("\n")}`);
      }
      checked = false;
      const written = writtenVisuals(pendingFiles(r));
      const text = await saveDraft(r, save);
      await showSaved(report(), written);
      onSaved?.();
      return [model, text].filter(Boolean).map((x) => `✓ ${x}`).join("\n") || undefined;
    },
  };
}
