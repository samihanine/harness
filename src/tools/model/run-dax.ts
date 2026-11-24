/** Runs a DAX query on a model (limits of Settings › Power BI: rows, time). */
import { z } from "zod";
import { dax, rowsText } from "@server/pbi/dax";
import { modelOf } from "../helpers";
import { tool } from "../tool";

export const runDax = tool({
  name: "run_dax",
  description:
    "Run a DAX query (EVALUATE …) on a semantic model to get figures, check values or explore. Prefer SUMMARIZECOLUMNS / aggregations to whole tables; rows are cut at the app's limit. Exact names only (read model_structure / search_model first).",
  args: z.object({ query: z.string().describe("A DAX query starting with EVALUATE (DEFINE allowed)."), dataset_id: z.string().optional() }),
  async run(a, ctx) {
    const d = modelOf(ctx, a.dataset_id);
    const rows = await dax({ id: d.id, groupId: d.groupId }, a.query);
    return rowsText(rows, 80);
  },
});
