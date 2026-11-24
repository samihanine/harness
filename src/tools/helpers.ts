/** What the tools of several screens share: the report and the models of the screen. */
import type { DatasetRef, ReportRef } from "@shared/types";
import type { ToolContext } from "./tool";

export function reportOf(ctx: ToolContext): ReportRef {
  if (!ctx.scope.report) throw new Error("No report is open on this screen: ask the user to choose one.");
  return ctx.scope.report;
}

/** The models the screen works with: its model, its report's model, the Data Scientist's models. */
export function modelsOf(ctx: ToolContext): DatasetRef[] {
  const s = ctx.scope;
  const list: DatasetRef[] = [];
  if (s.dataset) list.push(s.dataset);
  if (s.report?.datasetId) list.push({ id: s.report.datasetId, groupId: s.report.datasetGroupId, name: `model of ${s.report.name}` });
  for (const d of s.datasets ?? []) list.push(d);
  return list.filter((d, i) => list.findIndex((x) => x.id === d.id) === i);
}

/** The model a tool call is about: the one named (id or name), else the only one. */
export function modelOf(ctx: ToolContext, idOrName?: string): DatasetRef {
  const list = modelsOf(ctx);
  if (!list.length) throw new Error("No semantic model on this screen.");
  if (!idOrName) {
    if (list.length > 1) throw new Error(`Several models: give dataset_id, one of ${list.map((d) => `${d.id} ("${d.name}")`).join(", ")}.`);
    return list[0];
  }
  const hit = list.find((d) => d.id === idOrName || d.name.toLowerCase() === idOrName.toLowerCase());
  if (!hit) {
    if (list.length === 1) return list[0];
    throw new Error(`No model "${idOrName}". Models: ${list.map((d) => `${d.id} ("${d.name}")`).join(", ")}.`);
  }
  return hit;
}

/** The document the screen edits (the browser's copy), changed and sent back to the screen. */
export const docOf = <T>(ctx: ToolContext) => ctx.scope.doc as T | undefined;
export function setDoc(ctx: ToolContext, doc: unknown) {
  ctx.scope.doc = doc;
  ctx.effect("doc", doc);
}
