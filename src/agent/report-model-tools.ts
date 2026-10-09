/** Report builder: report measures (no model write needed) and parameters (model write needed). */
import { z } from "zod";
import { refreshDatasetText } from "@/lib/library";
import { addFieldParameter, addNumericParameter } from "@/lib/parameters";
import { layoutIssues, type Report } from "@/lib/pbi";
import { addShape, deleteReportMeasure, reportMeasuresText, setInteractions, setPositions, setSlicerSelection, setTextbox, upsertReportMeasure } from "@/lib/report-def";
import type { DatasetEntry, ReportEntry } from "@/lib/store";
import { refreshModel } from "@/server/ms";
import { tool, type Tool } from "./loop";

/** The user's pending edits are saved first: the definition is then rewritten and the editor reloaded. */
export async function rewrite(r: Report, task: () => Promise<string>) {
  await r.save().catch(() => undefined);
  const text = await task();
  // The editor reloads the new definition: the next tools wait until it is ready.
  const ready = new Promise<void>((resolve) => {
    const done = () => (r.off("loaded"), resolve());
    r.on("loaded", done);
    setTimeout(done, 20_000);
  });
  await r.reload().catch(() => undefined);
  await ready;
  return text;
}

export function reportModelTools({ report, entry, dataset, changed }: {
  report: () => Report | undefined;
  entry: () => ReportEntry | undefined;
  dataset: () => DatasetEntry | undefined;
  /** Called after a change (report measures reloaded by the page). */
  changed: () => unknown;
}): Tool[] {
  const need = () => {
    const r = report(), e = entry(), d = dataset();
    if (!r || !e || !d) throw new Error("Open a report first.");
    return { r, e, d, reportRef: { id: e.id, groupId: e.groupId }, datasetRef: { id: d.id, groupId: d.groupId } };
  };
  const around = async (task: () => Promise<string>) => {
    const text = await rewrite(need().r, task);
    await changed();
    return text;
  };
  const visualByName = async (name: string) => {
    const { r } = need();
    const page = (await r.getPages()).find((p) => p.isActive)!;
    const visuals = await page.getVisuals();
    const norm = (s = "") => s.toLowerCase().replace(/[\s_-]+/g, " ").trim();
    const v = visuals.find((x) => x.name === name) ?? visuals.find((x) => norm(x.title) === norm(name));
    if (!v) throw new Error(`Visual "${name}" not on the current page. Visuals: ${visuals.map((x) => `"${x.title ?? ""}" (name=${x.name})`).join(", ")}`);
    return { page: page.name, visual: v.name, all: visuals };
  };
  const activeVisuals = async () => (await (await need().r.getPages()).find((p) => p.isActive)!.getVisuals());
  /** Real positions after a change (read back from the editor), plus overlaps. `names` empty = every visual. */
  const layoutNow = async (names: string[]) => {
    const page = (await need().r.getPages()).find((p) => p.isActive)!;
    const visuals = await page.getVisuals();
    const shown = visuals.filter((v) => !names.length || names.includes(v.name));
    const pos = (v: (typeof visuals)[number]) => `x=${Math.round(v.layout.x ?? 0)} y=${Math.round(v.layout.y ?? 0)} w=${Math.round(v.layout.width ?? 0)} h=${Math.round(v.layout.height ?? 0)} z=${Math.round(v.layout.z ?? 0)}`;
    return ["Positions now:", ...shown.map((v) => `- ${v.type} "${v.title ?? ""}" name=${v.name} ${pos(v)}`), ...layoutIssues(visuals, page.defaultSize)].join("\n");
  };
  const writable = () => {
    const { d } = need();
    if (d.info?.source !== "tmdl")
      throw new Error(
        "Parameters are tables of the semantic model and you cannot edit this model. Alternative without model access: a report measure reading a slicer (e.g. SELECTEDVALUE('Table'[Column])), or ask the model owner.",
      );
  };
  const refresh = async () => {
    const { d, datasetRef } = need();
    const r = await refreshModel({ data: { datasetId: datasetRef.id, groupId: datasetRef.groupId } }).catch((e) => ({ status: "Failed", error: String(e) }));
    await refreshDatasetText(d).catch(() => undefined);
    return `model refresh ${r.status}${"error" in r && r.error ? `: ${r.error}` : ""}`;
  };
  return [
    tool({
      name: "list_report_measures",
      description: "Lists the report measures (measures stored in this report, not in the model).",
      args: z.object({}),
      readOnly: true,
      run: async () => reportMeasuresText(need().reportRef),
    }),
    tool({
      name: "add_report_measure",
      description:
        "Creates or replaces a report measure (DAX), usable by the visuals of this report only. No write access to the model needed. The DAX is checked first; errors come back. Use it for KPIs, selections (SELECTEDVALUE of a slicer column), titles, conditional colors.",
      args: z.object({
        table: z.string().describe("Model table that holds it (home table)"),
        name: z.string(),
        expression: z.string().describe("DAX expression (without 'Name =')"),
        format: z.string().optional().describe('Format string, e.g. "#,0", "0.0%", "$#,0"'),
        folder: z.string().optional(),
        description: z.string().optional(),
      }),
      run: (m) => {
        const { reportRef, datasetRef, d } = need();
        return around(() => upsertReportMeasure(reportRef, datasetRef, d.info, m));
      },
    }),
    tool({
      name: "delete_report_measure",
      description: "Deletes a report measure.",
      args: z.object({ name: z.string() }),
      run: ({ name }) => around(() => deleteReportMeasure(need().reportRef, name)),
    }),
    tool({
      name: "set_interactions",
      description:
        'Sets how a selection / click in a visual acts on the other visuals of the page ("Edit interactions"): "filter", "highlight", "none" or "default". E.g. clicking a row of a table filters only the detail card: source = the table, targets = detail → filter, every other visual → none.',
      args: z.object({
        source: z.string().describe("Visual name (preferred) or title"),
        targets: z.array(z.object({ visual: z.string(), type: z.enum(["filter", "highlight", "none", "default"]) })),
      }),
      run: ({ source, targets }) =>
        around(async () => {
          const s = await visualByName(source);
          const resolved = await Promise.all(targets.map(async (t) => ({ visual: (await visualByName(t.visual)).visual, type: t.type })));
          return setInteractions(need().reportRef, s.page, s.visual, resolved);
        }),
    }),
    tool({
      name: "set_textbox",
      description: "Writes the text of a text box (create it first with create_visual type textbox): size in points, bold, color, alignment. Use it for report titles.",
      args: z.object({
        visual: z.string(),
        text: z.string().describe("Text; \\n for a new paragraph"),
        size: z.number().optional().describe("Font size in points, e.g. 24 for a title"),
        bold: z.boolean().optional(),
        color: z.string().optional(),
        align: z.enum(["left", "center", "right"]).optional(),
      }),
      run: ({ visual, ...t }) =>
        around(async () => {
          const v = await visualByName(visual);
          return setTextbox(need().reportRef, v.page, v.visual, t);
        }),
    }),
    tool({
      name: "arrange_visuals",
      description:
        "Moves / resizes one or several visuals of the current page at once (x, y, width, height; page size in the context). Returns the real positions afterwards and any overlap: check them before telling the user.",
      args: z.object({
        moves: z.array(z.object({ visual: z.string(), x: z.number().optional(), y: z.number().optional(), width: z.number().optional(), height: z.number().optional() })),
      }),
      run: async ({ moves }) => {
        const resolved = await Promise.all(moves.map(async (m) => ({ ...m, visual: (await visualByName(m.visual)).visual })));
        const page = (await visualByName(resolved[0].visual)).page;
        await around(() => setPositions(need().reportRef, page, resolved));
        return layoutNow(resolved.map((m) => m.visual));
      },
    }),
    tool({
      name: "add_shape",
      description:
        'Adds a colored shape to the current page, e.g. a light panel BEHIND some visuals ("a rectangle behind the table and the card"): around = those visuals (the box is computed with a margin), behind = true. Or give x, y, width, height.',
      args: z.object({
        color: z.string().describe("Hex color, e.g. #FCE4D6"),
        around: z.array(z.string()).optional().describe("Visuals the shape surrounds"),
        margin: z.number().optional().describe("Space around those visuals, default 10"),
        x: z.number().optional(),
        y: z.number().optional(),
        width: z.number().optional(),
        height: z.number().optional(),
        shape: z.enum(["rectangle", "roundedRectangle", "oval"]).optional(),
        behind: z.boolean().optional().describe("Under the other visuals (default true)"),
        transparency: z.number().optional().describe("0-100"),
      }),
      run: async ({ color, around: wrapped, margin = 10, shape, behind = true, transparency, ...box }) => {
        const page = (await need().r.getPages()).find((p) => p.isActive)!.name;
        let b = box as { x?: number; y?: number; width?: number; height?: number };
        if (wrapped?.length) {
          const all = await activeVisuals();
          const names = await Promise.all(wrapped.map(async (w) => (await visualByName(w)).visual));
          const vs = all.filter((v) => names.includes(v.name));
          const x = Math.min(...vs.map((v) => v.layout.x ?? 0)) - margin, y = Math.min(...vs.map((v) => v.layout.y ?? 0)) - margin;
          const r = Math.max(...vs.map((v) => (v.layout.x ?? 0) + (v.layout.width ?? 0))) + margin, bottom = Math.max(...vs.map((v) => (v.layout.y ?? 0) + (v.layout.height ?? 0))) + margin;
          b = { x: Math.max(0, x), y: Math.max(0, y), width: r - Math.max(0, x), height: bottom - Math.max(0, y) };
        }
        if (b.x === undefined || b.y === undefined || !b.width || !b.height) throw new Error("Give around (visuals) or x, y, width and height.");
        const text = await around(() => addShape(need().reportRef, page, { ...(b as Required<typeof b>), color, shape, behind, transparency }));
        return `${text}\n${await layoutNow([])}`;
      },
    }),
    tool({
      name: "set_slicer_selection",
      description: 'Slicer settings: single or multiple selection, "Select all" option, and style: "dropdown" (a select), "list" or "tile" (buttons).',
      args: z.object({ visual: z.string(), single: z.boolean().optional(), select_all: z.boolean().optional(), style: z.enum(["dropdown", "list", "tile"]).optional() }),
      run: ({ visual, single, select_all, style }) =>
        around(async () => {
          const v = await visualByName(visual);
          return setSlicerSelection(need().reportRef, v.page, v.visual, { single, selectAll: select_all, style });
        }),
    }),
    tool({
      name: "add_field_parameter",
      description: "Adds a field parameter to the model: a slicer that switches BETWEEN SEVERAL measures or columns in the visuals using it (useless with one field). Needs write access to the model; refreshes it.",
      args: z.object({ name: z.string(), fields: z.array(z.object({ label: z.string(), field: z.string().describe("Table[Measure] or Table[Column]") })) }),
      run: ({ name, fields }) =>
        around(async () => {
          writable();
          await addFieldParameter(need().datasetRef, name, fields);
          return `Field parameter '${name}' added (use '${name}'[${name}] in a slicer and in the visuals' fields); ${await refresh()}.`;
        }),
    }),
    tool({
      name: "add_numeric_parameter",
      description: "Adds a numeric (what-if) parameter to the model: a slider table and the measure '<name> Value'. Needs write access to the model; refreshes it.",
      args: z.object({ name: z.string(), min: z.number(), max: z.number(), step: z.number(), value: z.number().describe("Default value") }),
      run: ({ name, ...p }) =>
        around(async () => {
          writable();
          await addNumericParameter(need().datasetRef, name, p);
          return `Numeric parameter '${name}' added: slicer on '${name}'[${name}], value in [${name} Value]; ${await refresh()}.`;
        }),
    }),
  ];
}
