/** Changes a visual's fields, type, title, sort or Top N, keeping its place and formatting. */
import { z } from "zod";
import { draftModel, pageOf, reportMeasures, visualOf, visualSchema, writeJson } from "@server/pbi/draft";
import { buildVisual } from "@server/pbi/visual-spec";
import { reportOf } from "../helpers";
import { tool } from "../tool";
import { draft } from "./shared";
import { visualArgs } from "./add-visual";

const fieldsBack = (q: any) => Object.fromEntries(Object.entries<any>(q?.queryState ?? {}).map(([role, s]) => [role === "Rows" ? "Rows" : role, (s.projections ?? []).map((p: any) => p.queryRef)]));

export const updateVisual = tool({
  name: "update_visual",
  description: "Change a visual of the report being built: its fields (all of them, by role), its type, title, sort or Top N. Its position and formatting are kept. Give only what changes.",
  args: z.object({ page: z.string(), visual: z.string().describe("Name or title."), type: z.string().optional(), ...visualArgs }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const page = pageOf(d, a.page);
    const { path, json: old } = visualOf(d, page.name, a.visual);
    const info = await draftModel(reportOf(ctx));
    const type = a.type ?? old.visual?.visualType;
    const fields = a.fields ?? Object.fromEntries(Object.entries(fieldsBack(old.visual?.query)).map(([r, f]) => [r, (f as string[]).map((x) => x.replace(/^(\w+)\((.+)\.(.+)\)$/, "$1('$2'[$3])").replace(/^([^(]+)\.([^.]+)$/, "'$1'[$2]"))]));
    const title = a.title ?? (String(old.visual?.visualContainerObjects?.title?.[0]?.properties?.text?.expr?.Literal?.Value ?? "").replace(/^'(.*)'$/s, "$1") || undefined);
    const built = buildVisual({ type, x: old.position.x, y: old.position.y, width: old.position.width, height: old.position.height, fields, title, sort: a.sort, top: a.top, text: a.text }, old.name, old.position.z ?? 0, info, reportMeasures(d), old.$schema ?? visualSchema(d));
    // Formatting kept: objects and container objects of the old visual, the new title over them.
    built.visual = { ...built.visual, objects: { ...(old.visual?.objects ?? {}), ...((built.visual as any).objects ?? {}) }, visualContainerObjects: { ...(old.visual?.visualContainerObjects ?? {}), ...((built.visual as any).visualContainerObjects ?? {}) } };
    built.position = { ...old.position };
    if (!a.top && old.filterConfig) (built as any).filterConfig = old.filterConfig;
    writeJson(d, path, built);
    return `Visual ${old.name} updated (${type}).`;
  },
});
