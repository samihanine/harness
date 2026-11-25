/**
 * What the user can do with a report: "edit" when its PBIR definition is readable (the Builder writes it),
 * "editor" when it is editable but only in another format (PBIR-Legacy: the Builder works in the embedded
 * Power BI editor), "view" when the API gives no definition. Kept a while (one Fabric call per report).
 */
import type { Ref, ReportAccess } from "@shared/types";
import { anyReportDefinition, reportDefinition } from "./definitions";

const cache = new Map<string, { at: number; access: Promise<ReportAccess> }>();

export function reportAccess(r: Ref): Promise<ReportAccess> {
  if (r.appId) return Promise.resolve("view");
  const hit = cache.get(r.id);
  if (hit && Date.now() - hit.at < 30 * 60_000) return hit.access;
  const access = reportDefinition(r).then(
    () => "edit" as const,
    () => anyReportDefinition(r).then(() => "editor" as const, () => "view" as const),
  );
  cache.set(r.id, { at: Date.now(), access });
  return access;
}

/** Several reports at once, a few at a time (each is a Fabric call). */
export async function reportsAccess(refs: Ref[]) {
  const out: Record<string, ReportAccess> = {};
  for (let i = 0; i < refs.length; i += 6) await Promise.all(refs.slice(i, i + 6).map(async (r) => (out[r.id] = await reportAccess(r))));
  return out;
}
