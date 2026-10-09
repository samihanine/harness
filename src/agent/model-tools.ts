/** Model tools: the AI gets a summary of the semantic model, and reads the details it needs. */
import { z } from "zod";
import { describeMeasure, describeTable, searchModel } from "@/lib/model";
import { all, type DatasetEntry } from "@/lib/store";
import { tool, type Tool } from "./loop";

export function modelTools(dataset: () => DatasetEntry | undefined): Tool[] {
  const info = () => {
    const d = dataset();
    if (!d?.info) throw new Error("The model structure is not read yet: use \"Read again\" on the model in the Library.");
    return d.info;
  };
  return [
    tool({
      name: "describe_table",
      description: "Details of a table of the model: columns (type, format, description, sample values or range), calculated column formulas, hierarchies, measures with their DAX, relationships.",
      args: z.object({ table: z.string() }),
      readOnly: true,
      run: async ({ table }) => describeTable(info(), table),
    }),
    tool({
      name: "get_measure",
      description: "A measure: its DAX, format, description, the measures it uses and that use it, the library reports showing it.",
      args: z.object({ measure: z.string().describe("Measure name, with or without brackets") }),
      readOnly: true,
      run: async ({ measure }) => describeMeasure(info(), measure, await all("reports")),
    }),
    tool({
      name: "search_model",
      description: "Finds tables, columns and measures whose name, description, values or DAX contain a text (e.g. a business term or a value like 'Finalize').",
      args: z.object({ query: z.string() }),
      readOnly: true,
      run: async ({ query }) => searchModel(info(), query),
    }),
  ];
}
