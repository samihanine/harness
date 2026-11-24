/** A new visual from a short spec: the PBIR is written by code, its fields checked against the model. */
import { z } from "zod";
import { draftModel, pageOf, reportMeasures, visualSchema, visualsOf, writeJson } from "@server/pbi/draft";
import { buildVisual, newName, ROLES, TYPE_NOTES } from "@server/pbi/visual-spec";
import { reportOf } from "../helpers";
import { tool } from "../tool";
import { draft } from "./shared";

export const visualArgs = {
  fields: z.record(z.string(), z.array(z.string())).optional().describe("Fields by role, e.g. {\"Category\":[\"'Date'[Month]\"],\"Y\":[\"[Revenue]\"]}. A column can be aggregated: Sum('Sales'[Amount])."),
  title: z.string().optional(),
  sort: z.object({ field: z.string(), descending: z.boolean().optional() }).optional(),
  top: z.object({ count: z.number(), by: z.string(), bottom: z.boolean().optional() }).optional().describe("Keep the first N items by a value."),
  text: z.string().optional().describe("textbox: its text."),
};

export const addVisual = tool({
  name: "add_visual",
  description: `Add a visual to a page of the report being built. Types and roles: ${Object.entries(ROLES).map(([t, r]) => `${t} (${r.join(", ") || "no field"})`).join("; ")}. ${TYPE_NOTES}. Place visuals afterwards with layout_page (or give x, y, width, height).`,
  args: z.object({
    page: z.string(),
    type: z.string(),
    ...visualArgs,
    x: z.number().optional(),
    y: z.number().optional(),
    width: z.number().optional(),
    height: z.number().optional(),
  }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const page = pageOf(d, a.page);
    const info = await draftModel(reportOf(ctx));
    const name = newName();
    const z = visualsOf(d, page.name).length * 1000;
    const json = buildVisual({ type: a.type, x: a.x ?? 0, y: a.y ?? 0, width: a.width ?? 400, height: a.height ?? 250, fields: a.fields, title: a.title, sort: a.sort, top: a.top, text: a.text }, name, z, info, reportMeasures(d), visualSchema(d));
    writeJson(d, `definition/pages/${page.name}/visuals/${name}/visual.json`, json);
    return `Visual ${a.type}${a.title ? ` "${a.title}"` : ""} added on "${page.displayName}": name=${name}.`;
  },
});
