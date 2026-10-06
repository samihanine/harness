/** Power BI embedding with the user's own (AAD) token. */
import * as pbi from "powerbi-client";
import "powerbi-report-authoring";
import type { ReportRef } from "./api";
import { embedUrl } from "./api";
import { getToken } from "./auth";

export { pbi };

const service = new pbi.service.Service(pbi.factories.hpmFactory, pbi.factories.wpmpFactory, pbi.factories.routerFactory);

const loaded = (embed: pbi.Embed) =>
  new Promise<void>((resolve, reject) => {
    embed.on("loaded", () => resolve());
    embed.on("error", (event) => {
      const detail = event.detail as { message?: string; detailedMessage?: string } | undefined;
      reject(new Error(detail?.detailedMessage ?? detail?.message ?? "Could not load the report"));
    });
  });

export async function embedReport(element: HTMLElement, ref: ReportRef, { edit = false, pageName }: { edit?: boolean; pageName?: string } = {}) {
  service.reset(element);
  const report = service.embed(element, {
    type: "report",
    id: ref.reportId,
    embedUrl: embedUrl(ref),
    accessToken: await getToken(),
    tokenType: pbi.models.TokenType.Aad,
    permissions: edit ? pbi.models.Permissions.ReadWrite : pbi.models.Permissions.Read,
    viewMode: pbi.models.ViewMode.View,
    pageName,
    settings: {
      panes: { filters: { visible: false, expanded: false }, pageNavigation: { visible: false } },
      background: pbi.models.BackgroundType.Transparent,
    },
  }) as pbi.Report;
  await loaded(report);
  return report;
}

/** One visual of a report. */
export async function embedVisual(element: HTMLElement, ref: ReportRef, pageName: string, visualName: string) {
  service.reset(element);
  const visual = service.embed(element, {
    type: "visual",
    id: ref.reportId,
    embedUrl: embedUrl(ref),
    accessToken: await getToken(),
    tokenType: pbi.models.TokenType.Aad,
    pageName,
    visualName,
    settings: { background: pbi.models.BackgroundType.Transparent },
  } as pbi.IVisualEmbedConfiguration) as pbi.Visual;
  await loaded(visual);
  return visual;
}

export const resetEmbed = (element: HTMLElement) => service.reset(element);
