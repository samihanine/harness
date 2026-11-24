/** Measures of a model written in its definition (TMDL through Fabric), then checked with DAX. */
import { z } from "zod";
import { readModelFiles, writeModelFiles } from "@server/pbi/definitions";
import { deleteMeasure as remove, tableFile, upsertMeasure } from "@server/pbi/tmdl-edit";
import { dax, rowsText } from "@server/pbi/dax";
import { forgetModel } from "@server/pbi/models";
import { modelOf } from "../helpers";
import { tool } from "../tool";

export const saveMeasure = tool({
  name: "save_measure",
  description: "Create or replace a measure in a table of the model (written to the model at once, then evaluated to check it). Test the DAX with run_dax first.",
  args: z.object({ table: z.string(), name: z.string(), expression: z.string().describe("DAX, without 'Name ='."), format: z.string().optional().describe('e.g. "#,0", "0.0%", "\\$#,0.00"'), folder: z.string().optional(), description: z.string().optional(), dataset_id: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const d = modelOf(ctx, a.dataset_id);
    if (d.access === "view") throw new Error("This model is view only.");
    const files = await readModelFiles(d.id, d.groupId);
    const f = tableFile(files, a.table);
    const changed = upsertMeasure(f, { name: a.name, expression: a.expression, format: a.format, folder: a.folder, description: a.description });
    ctx.status("Writing the model…");
    await writeModelFiles(d.id, files.map((x) => (x.path === f.path ? changed : x)), d.groupId);
    forgetModel(d.id);
    ctx.effect("model");
    const check = await dax(d, `EVALUATE ROW("value", [${a.name.replace(/]/g, "]]")}])`).then((r) => rowsText(r), (e) => `⚠ evaluation failed: ${e instanceof Error ? e.message : e}`);
    return `Measure [${a.name}] saved in '${a.table}'. Value over all the data:\n${check}`;
  },
});

export const deleteMeasure = tool({
  name: "delete_measure",
  description: "Delete a measure of the model (check with search_model that no other measure uses it).",
  args: z.object({ table: z.string(), name: z.string(), dataset_id: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const d = modelOf(ctx, a.dataset_id);
    const files = await readModelFiles(d.id, d.groupId);
    const f = tableFile(files, a.table);
    await writeModelFiles(d.id, files.map((x) => (x.path === f.path ? remove(f, a.name) : x)), d.groupId);
    forgetModel(d.id);
    ctx.effect("model");
    return `Measure [${a.name}] deleted.`;
  },
});
