/** The structure of a semantic model: tables, columns (with value examples), measures, relationships. */
import { z } from "zod";
import { modelInfo } from "@server/pbi/models";
import { modelSummary } from "@server/pbi/model";
import { modelOf } from "../helpers";
import { tool } from "../tool";

export const modelStructure = tool({
  name: "model_structure",
  description: "The structure of a semantic model: its tables with their columns, measures by folder and relationships (exact names to use in DAX and visuals). Read it before writing DAX or choosing fields.",
  args: z.object({ dataset_id: z.string().optional().describe("Only when the screen has several models."), focus: z.string().optional().describe("Words of the request: a big model then details the tables they are about.") }),
  async run(a, ctx) {
    return modelSummary(await modelInfo(modelOf(ctx, a.dataset_id)), a.focus);
  },
});
