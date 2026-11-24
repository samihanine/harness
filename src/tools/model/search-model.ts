/** Finds tables, columns, measures by a word (names, descriptions, values, DAX). */
import { z } from "zod";
import { modelInfo } from "@server/pbi/models";
import { searchModel as search } from "@server/pbi/model";
import { modelOf } from "../helpers";
import { tool } from "../tool";

export const searchModel = tool({
  name: "search_model",
  description: "Search a model for a word: tables, columns (also by their values, e.g. a country name), measures (also inside their DAX). Use it to map the user's words to the model's exact names.",
  args: z.object({ query: z.string(), dataset_id: z.string().optional() }),
  async run(a, ctx) {
    return search(await modelInfo(modelOf(ctx, a.dataset_id)), a.query);
  },
});
