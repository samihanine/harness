/** One table of a model in detail: columns (type, format, value examples), measures with their DAX, relationships. */
import { z } from "zod";
import { modelInfo } from "@server/pbi/models";
import { describeMeasure, describeTable as describe } from "@server/pbi/model";
import { modelOf } from "../helpers";
import { tool } from "../tool";

export const describeTable = tool({
  name: "describe_table",
  description: "One table of the model in detail (columns with types and value examples, measures with their DAX, relationships), or one measure when `measure` is given.",
  args: z.object({ table: z.string().optional(), measure: z.string().optional(), dataset_id: z.string().optional() }),
  async run(a, ctx) {
    const info = await modelInfo(modelOf(ctx, a.dataset_id));
    if (a.measure) return describeMeasure(info, a.measure);
    if (!a.table) throw new Error("Give a table or a measure.");
    return describe(info, a.table);
  },
});
