/** Real Power BI preview: the generated .pbix is imported in My workspace, then embedded. */
import type { ModelContent } from "@/pbi/types";
import { writePbix } from "@/pbi/pbix";
import { buildLayout } from "./layout";
import type { BuilderReport } from "./spec";

export const pbixOf = (report: BuilderReport, model: ModelContent) => writePbix(buildLayout(report, model), report.datasetId);

export function downloadPbix(report: BuilderReport, model: ModelContent) {
  const url = URL.createObjectURL(new Blob([pbixOf(report, model)]));
  Object.assign(document.createElement("a"), { href: url, download: `${report.title || "report"}.pbix` }).click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

type Import = { id: string; importState: "Publishing" | "Succeeded" | "Failed"; reports?: { id: string }[]; error?: { code?: string; details?: string } };

/** Imports (or overwrites) the report in My workspace; resolves to its report id. */
export async function publishToMyWorkspace(report: BuilderReport, model: ModelContent, log: (step: string) => void = () => {}) {
  log("Uploading to My workspace…");
  const form = new FormData();
  form.append("file", new Blob([pbixOf(report, model)]), `${report.title}.pbix`);
  const name = encodeURIComponent(`[builder] ${report.title}.pbix`);
  const response = await fetch(`/api/pbi/v1.0/myorg/imports?datasetDisplayName=${name}&nameConflict=CreateOrOverwrite`, { method: "POST", body: form });
  const started = (await response.json().catch(() => ({}))) as Partial<Import> & { error?: { message?: string } };
  if (!response.ok || !started.id) throw new Error(started.error?.message ?? `Import refused (${response.status})`);
  for (let i = 0; i < 60; i++) {
    await new Promise((r) => setTimeout(r, 2000));
    const state = (await (await fetch(`/api/pbi/v1.0/myorg/imports/${started.id}`)).json()) as Import;
    log(`Import: ${state.importState}…`);
    if (state.importState === "Succeeded" && state.reports?.[0]) return state.reports[0].id;
    if (state.importState === "Failed") throw new Error(`Import failed: ${state.error?.code ?? ""} ${state.error?.details ?? ""}`);
  }
  throw new Error("Import timed out");
}
