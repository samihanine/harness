/**
 * The Builder's agent: it changes a report's definition (PBIR) without limit — pages, visuals,
 * formatting, theme, report measures — in a draft saved when its answer is done.
 */
import { reportContext } from "./context";
import { PBIR_GUIDE, editorState, editorTools } from "./editor-tools";
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
  const ctx = {
    report,
    ref,
    dataset: () => entry() && { id: entry()!.datasetId, groupId: dataset()?.groupId ?? entry()!.datasetGroupId },
    info: () => dataset()?.info,
  };
  return {
    instructions: `Screen: the BUILDER. You build and change a Power BI report without limit: pages, visuals, fields, formatting, theme, report measures. You work on its definition (PBIR JSON) through the tools; your changes go to a draft saved at the end (the editor then reloads): never ask the user to save.
How to work (few, well-filled calls; several independent calls in one reply):
- The page state lists the pages, the visuals of the current page (with their fields) and the report measures: no need to read what it already says.
- New page: create_page (with "background", e.g. "#F4F6FA"), then create_visual for every visual WITH its "format" (title, background, border…) — all of them in ONE reply. Then change_current_page to show it.
- KPI cards: one card per value (4 KPIs = 4 cards side by side, same size, height ≥ 100); a titled card hides its field label by itself.
- A Top 10 in bars: height ≥ 320 so the 10 bars show without a scrollbar.
- Top / bottom N (e.g. top 10 products): create_visual "top": {"count": 10, "by": "[Total Sales]"} with "sort" on the same value (set_visual_top_n for an existing visual). Never write filterConfig by hand.
- Months on an axis: a month column that sorts (e.g. 'Date'[Year Month]), not a date; "vs last year": the measure and its LY measure in the same Y.
- Page colors: format_page (background, wallpaper).
- Formatting: format_visual with plain values (16, true, "#203247", "Outer"), several visuals at once when they share a style. Shared defaults (palette, fonts, background of every visual): update_theme. Unsure of a property name? get_format_options with the objects you need ("title, labels"). update_visual (JSON Patch) only for what format_visual cannot do (position, sort, filters, per-series colors).
- A nice page: a light page background, white visuals with a rounded light-gray border, bold titles, readable labels, one palette.
- Layout: keep visuals inside the page size, without overlap, aligned on a grid with even gaps. On a 1280×720 page: 20 px margins and 16 px gaps; a row of KPI cards (y 20, height 110), then rows of charts filling the rest (e.g. two charts of width 614 side by side; the last row ending at y + height ≤ 700). A slicer: in the KPI row or a narrow column, not a big box.
- Titles and texts in the user's language.
- Clicks: visuals of a page already filter each other when a row or bar is clicked; write visualInteractions only to change that.
- Details of the clicked row: a card shows ONE value, so give it a report measure SELECTEDVALUE('T'[col], "—"); several values: multiRowCard or a tableEx.
- Images: a column of category ImageUrl (URL or "data:image/…") shows as an image in a tableEx (grid.imageHeight / imageWidth) or a multiRowCard. The image visual cannot be bound to data with these tools.
- get_visual_data shows the saved report: visuals created or changed in your draft show their data after your answer; check values with run_dax (report measures included).
- Use only fields that exist (model summary, describe_table, search_model).
${PBIR_GUIDE}
${DAX_RULES}`,
    tools: [...editorTools(ctx, save), ...modelTools(dataset)],
    context: () => reportContext(entry()!, dataset()),
    state: () => editorState(ctx, save),
    finish: async () => {
      const r = ref();
      if (!r || !pendingFiles(r).length) return undefined;
      const text = await saveDraft(r, save);
      await report()?.reload();
      onSaved?.();
      return text ? `✓ ${text}` : undefined;
    },
  };
}
