import { useRef } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { SparklesIcon } from "lucide-react";
import { Chat } from "@/components/chat";
import { ReportActions } from "@/components/report-actions";
import { Empty, PageHeader, SidePanel } from "@/components/side-panel";
import { Zoom } from "@/components/zoom";
import { AI_PANEL } from "@/config";
import { PBIR_GUIDE, editorState, editorTools } from "@/agent/editor-tools";
import type { Agent } from "@/agent/loop";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES } from "@/agent/pbi-tools";
import { useLast } from "@/lib/last";
import { pendingFiles, saveDraft } from "@/lib/report-draft";
import { useCollection } from "@/lib/store";
import { useEmbed } from "@/lib/use-embed";
import { ReportPicker, reportContext } from "./viewer";

export const Route = createFileRoute("/editor")({
  validateSearch: (s: Record<string, unknown>) => ({ report: (s.report as string) || undefined }),
  component: Editor,
});

/**
 * Report editor: the report in Power BI's own editor (authoring embed) and the AI, which changes its
 * definition (PBIR) without limit — pages, visuals, formatting, theme, report measures — in a draft
 * saved when its answer is done; the editor is then reloaded.
 */
function Editor() {
  const { report: reportId } = Route.useSearch();
  const navigate = useNavigate({ from: "/editor" });
  useLast("editor:report", reportId, (report) => navigate({ search: { report }, replace: true }));
  const entry = useCollection("reports").find((r) => r.id === reportId && r.access === "edit");
  const dataset = useCollection("datasets").find((d) => d.id === entry?.datasetId);
  const { element, report, error } = useEmbed(entry, "edit");
  // Read by the tools at call time (the agent object is made at each render).
  const live = useRef({ report, entry, dataset });
  live.current = { report, entry, dataset };
  const ref = () => live.current.entry && { id: live.current.entry.id, groupId: live.current.entry.groupId };
  const save = async () => live.current.report?.save();
  const ctx = {
    report: () => live.current.report,
    ref,
    dataset: () => live.current.entry && { id: live.current.entry.datasetId, groupId: live.current.dataset?.groupId ?? live.current.entry.datasetGroupId },
    info: () => live.current.dataset?.info,
  };

  const agent: Agent | undefined = entry && {
    instructions: `You build and change a Power BI report without limit: pages, visuals, fields, formatting, theme, report measures. You work on its definition (PBIR JSON) through the tools; your changes go to a draft saved when your answer is done (the editor then reloads): never ask the user to save.
How to work (few, well-filled calls; several independent calls in one reply):
- The page state lists the pages, the visuals of the current page (with their fields) and the report measures: no need to read what it already says.
- New page: create_page, then create_visual for each visual WITH its "format" (title, background, border, labels…) in the same call; the result needs no get_visual to check it. Then change_current_page to show it.
- Formatting: format_visual with plain values (16, true, "#203247", "Outer"), several visuals at once when they share a style. Shared defaults (palette, fonts, background of every visual): update_theme. Unsure of a property name? get_format_options with the objects you need ("title, labels"), once per visual type. update_visual (JSON Patch) only for what format_visual cannot do (position, sort, filters, per-series colors).
- A nice page: a light page background, white visuals with a rounded light-gray border, bold titles, readable labels, one palette.
- Layout: keep visuals inside the page size (see the page state), without overlap, aligned on a grid with even gaps.
- Use only fields that exist (model summary, describe_table, search_model); check values and totals with run_dax.
- Clicks: visuals of a page already filter each other when a row or bar is clicked (no setup needed); write visualInteractions only to change that.
- Details of the clicked row: a card shows ONE value, so give it a report measure SELECTEDVALUE('T'[col], "—") (shows "—" until a row is clicked); several values at once: multiRowCard or a tableEx.
- Images: a column of category ImageUrl (an URL or "data:image/…" value) shows as an image in a tableEx (set grid.imageHeight / imageWidth) or a multiRowCard. The image visual cannot be bound to data with these tools: never fake it by changing a visualType.
- get_visual_data shows the saved report: visuals created or changed in your draft show their data after your answer; check values with run_dax (report measures included).
- Decide, then do it once: do not create, undo and redo the same visuals. When a tool refuses a change, read why before trying another way.
- Keep each reply short JSON: split a long batch of operations over several replies rather than risk an invalid one.
- Do the whole request in one go, then answer with what you changed.
${PBIR_GUIDE}
${DAX_RULES}`,
    tools: [...editorTools(ctx, save), ...modelTools(() => live.current.dataset)],
    context: () => reportContext(entry, dataset),
    state: () => editorState(ctx, save),
    finish: async () => {
      const r = ref();
      const embed = live.current.report;
      if (!r || !embed || !pendingFiles(r).length) return undefined;
      const text = await saveDraft(r, save);
      await embed.reload();
      return text ? `✓ ${text}` : undefined;
    },
  };

  return (
    <div className="flex h-full">
      {AI_PANEL && <SidePanel tabs={[{ id: "ai", label: "AI assistant", icon: SparklesIcon, content: agent ? <Chat scope={`editor:${entry!.id}`} agent={agent} placeholder="Ask to add a page, build visuals, change the formatting…" /> : <Empty>Choose a report you can edit.</Empty> }]} />}
      <div className="flex min-w-0 flex-1 flex-col">
        <PageHeader
          screen="Report editor"
          picker={<ReportPicker editable value={reportId} onChange={(report) => navigate({ search: { report } })} />}
          actions={
            <>
              <ReportActions entry={entry} report={report} />
              <Zoom report={report} />
            </>
          }
        />
        {error && <p className="px-6 py-2 text-destructive">{error}</p>}
        {!entry && <Empty>Choose a report you can edit (its definition must be readable in PBIR format: see its access in the Library).</Empty>}
        <div ref={element} className={entry ? "min-h-0 flex-1" : "hidden"} />
      </div>
    </div>
  );
}
