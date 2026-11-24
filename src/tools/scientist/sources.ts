/** What the analysis can use: models, uploaded files (sheets and columns), tables already loaded. */
import { z } from "zod";
import { fileSheets, schemaText } from "@server/data/engine";
import { tool } from "../tool";
import { analysisOf, engine } from "./shared";

export const listSources = tool({
  name: "list_sources",
  description: "What the analysis can use: the semantic models (query them with DAX, then load_dax), the uploaded files with their sheets and columns (load_file), and the SQL tables and steps already loaded.",
  args: z.object({}),
  async run(_, ctx) {
    const s = ctx.scope;
    const files = await Promise.all(
      (s.files ?? []).map(async (f) =>
        fileSheets(f.id).then(
          (sheets) => `- file "${f.name}" (id ${f.id}): ${sheets.map((sh) => `sheet "${sh.name}" ${sh.rows.length} rows, columns ${sh.header.join(", ")}`).join("; ")}`,
          (e) => `- file "${f.name}" (id ${f.id}): not readable (${e instanceof Error ? e.message : e})`,
        ),
      ),
    );
    const a = analysisOf(ctx);
    return [
      `Semantic models: ${(s.datasets ?? []).map((d) => `"${d.name}" (id ${d.id})`).join(", ") || "none"}`,
      `Files:\n${files.join("\n") || "none"}`,
      `SQL database:\n${schemaText(await engine(ctx))}`,
      a.steps.length ? `Steps: ${a.steps.map((x) => x.name).join(", ")}` : "",
    ].filter(Boolean).join("\n\n");
  },
});
