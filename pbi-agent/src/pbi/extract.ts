/**
 * Content of a published report: the full definition from its .pbix when it can be downloaded,
 * otherwise what the embedding API exposes (pages, visuals, fields by role, filters, main format).
 */
import type { ReportRef } from "./api";
import { exportReport } from "./api";
import { embedReport, pbi, resetEmbed } from "./embed";
import { readPbix } from "./pbix";
import type { PageContent, ReportContent, VisualContent } from "./types";

const FORMAT = [
  ["title", "titleText"],
  ["title", "fontColor"],
  ["background", "color"],
  ["dataPoint", "defaultColor"],
  ["legend", "visible"],
  ["labels", "visible"],
] as const;

export async function extractReport(ref: ReportRef, log: (step: string) => void = () => {}) {
  log("Downloading the report definition…");
  try {
    return readPbix(await exportReport(ref));
  } catch (error) {
    log("Download not allowed: reading the report through embedding…");
    const content = await extractByEmbed(ref, log);
    content.warnings.unshift(`Full definition unavailable (${error instanceof Error ? error.message : error}): formatting is partial.`);
    return { content, connection: undefined };
  }
}

async function extractByEmbed(ref: ReportRef, log: (step: string) => void): Promise<ReportContent> {
  const element = document.createElement("div");
  element.style.cssText = "position:fixed;left:-10000px;top:0;width:1280px;height:720px;";
  document.body.append(element);
  const warnings: string[] = [];
  try {
    const report = await embedReport(element, ref, { edit: true }).catch(() => embedReport(element, ref));
    const editable = await report.switchMode(pbi.models.ViewMode.Edit).then(
      () => true,
      () => false,
    );
    if (!editable) warnings.push("No edit rights: fields by role and formatting may be missing.");
    const pages: PageContent[] = [];
    for (const page of await report.getPages()) {
      log(`Page ${page.displayName}…`);
      await page.setActive().catch(() => undefined);
      const visuals: VisualContent[] = [];
      for (const visual of await page.getVisuals()) visuals.push(await readVisual(visual));
      pages.push({
        name: page.name,
        displayName: page.displayName,
        width: page.defaultSize?.width ?? 1280,
        height: page.defaultSize?.height ?? 720,
        hidden: page.visibility === 1 || undefined,
        filters: await page.getFilters().catch(() => []),
        visuals,
      });
    }
    return { source: "embed", pages, filters: await report.getFilters().catch(() => []), warnings, extractedAt: new Date().toISOString() };
  } finally {
    resetEmbed(element);
    element.remove();
  }
}

const target = (t: Record<string, unknown>) =>
  t.measure ? `[${t.measure}]` : `${t.table}[${t.column ?? t.hierarchyLevel ?? t.hierarchy ?? ""}]`;

async function readVisual(visual: pbi.VisualDescriptor): Promise<VisualContent> {
  const fields: Record<string, string[]> = {};
  const capabilities = await visual.getCapabilities().catch(() => null);
  for (const role of capabilities?.dataRoles ?? [])
    fields[role.name] = (await visual.getDataFields(role.name).catch(() => [])).map((t) => target(t as unknown as Record<string, unknown>));
  const objects: Record<string, Record<string, unknown>> = {};
  for (const [objectName, propertyName] of FORMAT) {
    const property = await visual.getProperty({ objectName, propertyName }).catch(() => null);
    if (property?.value !== undefined && property.value !== null) (objects[objectName] ??= {})[propertyName] = property.value;
  }
  const { layout } = visual;
  return {
    name: visual.name,
    type: visual.type,
    title: visual.title,
    x: Math.round(layout.x ?? 0),
    y: Math.round(layout.y ?? 0),
    z: layout.z,
    width: Math.round(layout.width ?? 0),
    height: Math.round(layout.height ?? 0),
    hidden: layout.displayState?.mode === 1 || undefined,
    fields,
    filters: await visual.getFilters().catch(() => []),
    objects,
  };
}
