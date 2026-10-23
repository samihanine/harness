/**
 * The Builder's agent: it changes a report's definition (PBIR) without limit — pages, visuals,
 * formatting, theme, report measures — in a draft saved when its answer is done.
 */
import { reportContext } from "./context";
import { examplesText, referenceReports } from "@/lib/report-schema";
import { PBIR_GUIDE, draftProblems, editorState, editorTools } from "./editor-tools";
import { draftInfo, fieldParameterTool, saveModelDraft } from "./model-editor-tools";
import type { Agent } from "./loop";
import { modelTools } from "./model-tools";
import { DAX_RULES } from "./pbi-tools";
import type { Report } from "@/lib/pbi";
import { pendingFiles, saveDraft } from "@/lib/report-draft";
import type { DatasetEntry, ReportEntry } from "@/lib/store";

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
    instructions: `Screen: the BUILDER. You build and change a Power BI report without limit: pages, visuals, fields, formatting, theme, report measures. You work on its definition (PBIR JSON) through the tools; your changes go to a draft saved at the end (the editor then reloads): never ask the user to save.
How to work (few, well-filled calls; several independent calls in one reply):
- The page state lists the pages, the visuals of the current page (with their fields) and the report measures: no need to read what it already says.
- New page: create_page (with "background", e.g. "#F4F6FA"), then create_visual for every visual WITH its "format" (title, background, border…) and WITHOUT x / y / width / height — all of them in ONE reply; then ONE layout_page with the rows of the page; then change_current_page to show it.
- KPI cards: one card per value (4 KPIs = 4 cards side by side, same size, height ≥ 100); a titled card hides its field label by itself.
- A Top 10 in bars: height ≥ 320 so the 10 bars show without a scrollbar.
- Top / bottom N (e.g. top 10 products): create_visual "top": {"count": 10, "by": "[Total Sales]"} with "sort" on the same value (set_visual_top_n for an existing visual). Never write filterConfig by hand.
- Months on an axis: a month column that sorts (e.g. 'Date'[Year Month]), not a date; "vs last year": the measure and its LY measure in the same Y.
- Page colors: format_page (background, wallpaper).
- Formatting: format_visual with plain values (16, true, "#203247", "Outer"), several visuals at once when they share a style. Shared defaults (palette, fonts, background of every visual): update_theme. Unsure of a property name? get_format_options with the objects you need ("title, labels"). update_visual (JSON Patch) only for what format_visual cannot do (position, sort, filters, per-series colors).
- A nice page: a light page background, white visuals with a rounded light-gray border, bold titles, readable labels, one palette.
- Layout: layout_page by default (rows of visuals; the app computes the positions: margins, even gaps, no overlap), also to rearrange a page. A row of KPI cards (height 110) then rows of charts; a slicer in the KPI row or a narrow column (width weight 1 vs 3). Explicit x / y / width / height (create_visual, update_visual "/position") only when rows cannot say it (an overlay, a visual over another, a free placement the user asked for).
- Buttons (go to a page, back, a web link): create_button; a click action on any visual: set_visual_action.
- Conditional formatting (colors by value: a gradient, rules, or a measure returning a color): set_conditional_format. Target / threshold / average lines on a chart: set_reference_lines. Small multiples (one small chart per value): the "SmallMultiples" role of create_visual / set_visual_fields, the grid with format_visual {"smallMultiplesLayout": {"rowCount": 2, "columnCount": 3}}.
- Let the user switch the measure or the axis of visuals: a field parameter (create_field_parameter, when the model is editable), then 'Name'[Name] in a slicer and in the visuals' role. Its values show after the user refreshes the model.
- Reference reports (the "Reference reports" part of the context, chosen in the Reports tab): inspiration the user may refer to ("like the sales report", "the page Overview of …"). copy_report_page rebuilds one of their pages here (types, titles, fields, layout), even when their definition is closed; then format it.
- The draft is checked before it is saved (check_report: overlaps, out of the page, failing report measures): fix what it reports.
- Titles and texts in the user's language.
- Clicks: visuals of a page already filter each other when a row or bar is clicked; write visualInteractions only to change that.
- Details of the clicked row: a card shows ONE value, so give it a report measure SELECTEDVALUE('T'[col], "—"); several values: multiRowCard or a tableEx.
- Images: a column of category ImageUrl (URL or "data:image/…") shows as an image in a tableEx (grid.imageHeight / imageWidth) or a multiRowCard. The image visual cannot be bound to data with these tools.
- get_visual_data shows the saved report: visuals created or changed in your draft show their data after your answer; check values with run_dax (report measures included).
- Use only fields that exist (model summary, describe_table, search_model).
${PBIR_GUIDE}
${DAX_RULES}`,
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
      const text = await saveDraft(r, save);
      await report()?.reload();
      onSaved?.();
      return [model, text].filter(Boolean).map((x) => `✓ ${x}`).join("\n") || undefined;
    },
  };
}
