/**
 * Report builder: the AI reads and changes the report's definition (PBIR JSON) of the report, a page
 * or a visual, in a draft saved when its answer is done (see `finish` in the builder). Anything Power
 * BI stores can be changed this way, beyond what the authoring API offers.
 */
import { z } from "zod";
import { findVisual, type Report } from "@/lib/pbi";
import { discardDraft, patchDraftFile, pendingFiles, readDraftFile, undoLastSave, type PatchOp } from "@/lib/report-draft";
import type { ReportEntry } from "@/lib/store";
import { tool, type Tool } from "./loop";
import { rewrite } from "./report-model-tools";

const TARGET = z.string().describe('"report" (report.json: theme, report-wide settings), "page" (the current page: size, background), "page:<name>", or a visual of the current page (its name, or its title when no other visual has it)');

export function jsonTools({ report, entry }: { report: () => Report | undefined; entry: () => ReportEntry | undefined }): Tool[] {
  const need = () => {
    const r = report(), e = entry();
    if (!r || !e) throw new Error("Open a report first.");
    return { r, e, ref: { id: e.id, groupId: e.groupId }, save: () => r.save() };
  };
  /** The definition file of a target. */
  const fileOf = async (r: Report, target: string) => {
    const t = target.trim();
    if (t.startsWith("definition/")) return t;
    if (t === "report") return "definition/report.json";
    const pages = await r.getPages();
    const page = pages.find((p) => p.isActive) ?? pages[0];
    if (t === "page") return `definition/pages/${page.name}/page.json`;
    const named = t.match(/^page:(.+)$/)?.[1];
    if (named) {
      const p = pages.find((x) => x.name === named || x.displayName === named);
      if (!p) throw new Error(`No page "${named}". Pages: ${pages.map((x) => `${x.displayName} (${x.name})`).join(", ")}`);
      return `definition/pages/${p.name}/page.json`;
    }
    return `definition/pages/${page.name}/visuals/${findVisual(await page.getVisuals(), t).name}/visual.json`;
  };
  /** A file missing from the draft (e.g. a visual created after the draft was made) is found by making it again, when nothing is pending. */
  const withFresh = async <T,>(run: () => Promise<T>) => {
    try {
      return await run();
    } catch (e) {
      const { ref } = need();
      if (!/No file/.test(String(e)) || pendingFiles(ref).length) throw e;
      discardDraft(ref);
      return run();
    }
  };
  return [
    tool({
      name: "read_json",
      description: "Reads the JSON definition (PBIR) of the report, a page or a visual, as it is in your draft (your unsaved changes included): its whole formatting. Read before changing it.",
      args: z.object({ target: TARGET }),
      readOnly: true,
      run: async ({ target }) => {
        const { r, ref, save } = need();
        const file = await fileOf(r, target);
        return withFresh(() => readDraftFile(ref, save, file));
      },
    }),
    tool({
      name: "patch_json",
      description: `Changes the JSON definition of the report, a page or a visual with JSON Patch operations, in a draft: every change is saved automatically when you give your answer (do not call save_report for them). All operations of a call apply, or none.
- op: "add" (sets a key; also creates the missing objects on the way; "-" appends to an array), "replace" (an existing key), "remove".
- path: JSON Pointer from the file root, as read_json shows it, e.g. /visual/objects/legend/0/properties/position.
- Values are Power BI expressions: {"expr": {"Literal": {"Value": "'text'"}}} — text in single quotes inside the string, numbers with D ("14D"), "true" / "false"; a color: {"solid": {"color": {"expr": {"Literal": {"Value": "'#1F77B4'"}}}}}.
- Visual formatting: visual.objects (chart parts: legend, categoryAxis, valueAxis, labels, dataPoint…); visual.visualContainerObjects (title, subTitle, background, border, dropShadow, padding). Each is a list of {"properties": {…}} (dataPoint per series: add a "selector").
Example, the title: add /visual/visualContainerObjects/title/0/properties/text = {"expr": {"Literal": {"Value": "'Sales 2025'"}}}.`,
      args: z.object({
        target: TARGET,
        operations: z.array(z.object({ op: z.enum(["add", "replace", "remove"]), path: z.string(), value: z.unknown().optional() })),
      }),
      run: async ({ target, operations }) => {
        const { r, ref, save } = need();
        const file = await fileOf(r, target);
        return withFresh(() => patchDraftFile(ref, save, file, operations as PatchOp[]));
      },
    }),
    tool({
      name: "discard_json_changes",
      description: "Drops your unsaved JSON changes (the draft), e.g. after a wrong start.",
      args: z.object({}),
      run: async () => {
        const { ref } = need();
        const n = pendingFiles(ref).length;
        discardDraft(ref);
        return n ? `Unsaved changes to ${n} file(s) dropped.` : "Nothing was pending.";
      },
    }),
    tool({
      name: "undo_last_json_save",
      description: "Puts the report back as it was before your last saved JSON changes (only when the user asks to undo them).",
      args: z.object({}),
      run: async () => {
        const { r, ref } = need();
        return rewrite(r, () => undoLastSave(ref));
      },
    }),
  ];
}
