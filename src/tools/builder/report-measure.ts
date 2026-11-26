/**
 * Report measures: DAX measures kept in the report itself (its reportExtensions), not in the model — what a
 * user with only the Build permission on a model can add. Each is checked on the model with the report's other
 * measures before it is written.
 */
import { z } from "zod";
import { dax } from "@server/pbi/dax";
import { readJson, writeJson } from "@server/pbi/draft";
import { reportOf } from "../helpers";
import { tool } from "../tool";
import { draft } from "./shared";

type Measure = { name: string; dataType?: string; expression: string; formatString?: string };
type Ext = { name: string; entities: { name: string; measures: Measure[] }[] };
const PATH = "definition/reportExtensions.json";
const SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/reportExtension/1.0.0/schema.json";
const q = (s: string) => `'${s.replace(/'/g, "''")}'`;

export const reportMeasure = tool({
  name: "report_measure",
  description:
    "Add, change or remove a REPORT measure (kept in the report, not in the model: works with only the Build permission on the model). table: the model table it shows under; type text for a measure returning text (e.g. a color for conditional formatting); format e.g. \"#,0\", \"0.0%\". Checked on the model first. Use it in visuals as [Name].",
  args: z.object({ table: z.string(), name: z.string(), expression: z.string().optional(), type: z.enum(["number", "text"]).optional(), format: z.string().optional(), remove: z.boolean().optional() }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const ext: Ext = d.parts.some((p) => p.path === PATH) ? readJson(d, PATH) : { name: "extension", entities: [] };
    ext.entities ??= [];
    for (const e of ext.entities) e.measures = (e.measures ?? []).filter((m) => m.name.toLowerCase() !== a.name.toLowerCase());
    if (!a.remove) {
      if (!a.expression) throw new Error("Give the expression.");
      const r = reportOf(ctx);
      if (r.datasetId) {
        const others = ext.entities.flatMap((e) => e.measures.map((m) => `MEASURE ${q(e.name)}[${m.name}] = ${m.expression}`));
        await dax({ id: r.datasetId, groupId: r.datasetGroupId }, `DEFINE ${[...others, `MEASURE ${q(a.table)}[${a.name}] = ${a.expression}`].join("\n")}\nEVALUATE ROW("value", ${q(a.table)}[${a.name}])`).catch((e) => {
          throw new Error(`The measure does not work on the model: ${e instanceof Error ? e.message : e}`);
        });
      }
      let entity = ext.entities.find((e) => e.name.toLowerCase() === a.table.toLowerCase());
      if (!entity) ext.entities.push((entity = { name: a.table, measures: [] }));
      entity.measures.push({ name: a.name, dataType: a.type === "text" ? "Text" : "Double", expression: a.expression, ...(a.format ? { formatString: a.format } : {}) });
    }
    ext.entities = ext.entities.filter((e) => e.measures.length);
    writeJson(d, PATH, { $schema: SCHEMA, ...ext });
    return a.remove ? `Report measure [${a.name}] removed.` : `Report measure [${a.name}] saved in the report (use it as [${a.name}]).`;
  },
});
