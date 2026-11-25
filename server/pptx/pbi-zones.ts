/**
 * The Power BI views of a deck: the report found by its id or name, and a picture of the page or visual
 * with its filters (headless browser), kept 10 minutes so previews and exports do not take them again.
 */
import type { PbiZone } from "@shared/deck";
import { captureReport } from "../capture/browser";
import { reportCatalog } from "../pbi/catalog";
import { reportStructure } from "../pbi/report-structure";

const cache = new Map<string, { at: number; png: Promise<Buffer> }>();

export async function zoneReport(given: string) {
  const list = await reportCatalog();
  // A report's address from the browser works too.
  const report = given.match(/(?:reports\/|reportId=)([0-9a-f-]{36})/i)?.[1].toLowerCase() ?? given.trim();
  const r = list.find((x) => x.id === report) ?? list.find((x) => x.name.toLowerCase() === report.toLowerCase()) ?? list.find((x) => x.name.toLowerCase().includes(report.toLowerCase()));
  if (!r) throw new Error(`no report "${report}"`);
  return r;
}

const toFilters = (z: PbiZone) => z.filters.map((f) => ({ $schema: "http://powerbi.com/product/schema#basic", target: { table: f.table, column: f.column }, operator: "In", values: f.values, filterType: 1 }));

export function capturePbiZone(z: PbiZone): Promise<Buffer> {
  const key = JSON.stringify([z.report, z.page, z.visual, z.filters]);
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < 10 * 60_000) return hit.png;
  const png = (async () => {
    const r = await zoneReport(z.report);
    // The picture's proportions close to the zone's (a visual is cut from the page).
    const shot = await captureReport({ id: r.id, groupId: r.groupId, appId: r.appId }, { page: z.page, visual: z.visual, filters: toFilters(z), width: 1600, height: 900 });
    return shot.png;
  })();
  cache.set(key, { at: Date.now(), png });
  png.catch(() => cache.delete(key));
  return png;
}

/** The zones resolved for a live embed in the browser: the report's ids, the page and visual names. */
export async function resolveZones(zones: PbiZone[]) {
  return Promise.all(
    zones.map(async (z) => {
      try {
        const r = await zoneReport(z.report);
        const ref = { id: r.id, groupId: r.groupId, appId: r.appId };
        // The embed takes names, the deck may give titles.
        const s = (await reportStructure(ref).catch(() => undefined)) as { pages?: { name: string; displayName: string; visuals: { name: string; title?: string }[] }[] } | undefined;
        const page = z.page ? s?.pages?.find((p) => p.name === z.page || p.displayName.toLowerCase() === z.page!.toLowerCase()) : s?.pages?.[0];
        const visual = z.visual ? page?.visuals.find((v) => v.name === z.visual || (v.title ?? "").toLowerCase() === z.visual!.toLowerCase()) : undefined;
        return { ...z, ref, pageName: page?.name, visualName: visual?.name };
      } catch (e) {
        return { ...z, error: e instanceof Error ? e.message : String(e) };
      }
    }),
  );
}
