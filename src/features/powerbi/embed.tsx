/**
 * A Power BI report embedded in the page (powerbi-client), with the Power BI token from the server
 * (refreshed before it expires). The report object is what the browser actions act on (src/tools/browser).
 */
import { useEffect, useRef, useState } from "react";
import type * as PbiClient from "powerbi-client";
import type { Ref } from "@shared/types";
import { errorText, get } from "@/lib/api";

export type Report = PbiClient.Report;
let lib: Promise<{ pbi: typeof PbiClient; service: PbiClient.service.Service }> | undefined;
export const powerbi = () =>
  (lib ??= import("powerbi-client").then((pbi) => ({ pbi, service: new pbi.service.Service(pbi.factories.hpmFactory, pbi.factories.wpmpFactory, pbi.factories.routerFactory) })));

let token: { token: string; expiresAt: number } | undefined;
async function embedToken() {
  if (!token || token.expiresAt < Date.now() + 120_000) token = await get("/embed-token");
  return token!.token;
}

const embedUrl = (ref: Ref) => `https://app.powerbi.com/reportEmbed?reportId=${ref.id}${ref.appId ? `&appId=${ref.appId}` : ref.groupId ? `&groupId=${ref.groupId}` : ""}`;

/** Embeds a report in the element; `report` once it has loaded. `reloadKey` changing reloads it (after a save). */
export function useEmbed(ref: Ref | undefined, reloadKey = 0) {
  const element = useRef<HTMLDivElement>(null);
  const [report, setReport] = useState<Report>();
  const [error, setError] = useState<string>();
  const [loading, setLoading] = useState(false);
  useEffect(() => {
    const el = element.current;
    if (!ref || !el) return;
    let cancelled = false;
    setReport(undefined);
    setError(undefined);
    setLoading(true);
    let refresh: ReturnType<typeof setInterval> | undefined;
    void (async () => {
      try {
        const { pbi, service } = await powerbi();
        service.reset(el);
        const r = service.embed(el, {
          type: "report",
          id: ref.id,
          embedUrl: embedUrl(ref),
          accessToken: await embedToken(),
          tokenType: pbi.models.TokenType.Aad,
          permissions: pbi.models.Permissions.Read,
          viewMode: pbi.models.ViewMode.View,
          settings: { panes: { filters: { visible: true, expanded: false }, pageNavigation: { visible: true } }, background: pbi.models.BackgroundType.Transparent, layoutType: pbi.models.LayoutType.Custom, customLayout: { displayOption: pbi.models.DisplayOption.FitToPage } },
        }) as Report;
        r.on("loaded", async () => {
          if (cancelled) return;
          // The first view, to come back to (reset_view).
          (r as Report & { initialState?: string }).initialState = (await r.bookmarksManager.capture().catch(() => undefined))?.state;
          setReport(r);
          setLoading(false);
        });
        r.on("error", (e) => {
          const d = e.detail as { message?: string; detailedMessage?: string };
          if (!cancelled) (setError(d?.detailedMessage ?? d?.message ?? "Could not load the report"), setLoading(false));
        });
        refresh = setInterval(async () => {
          if (token && token.expiresAt < Date.now() + 300_000) await r.setAccessToken(await embedToken()).catch(() => undefined);
        }, 60_000);
      } catch (e) {
        if (!cancelled) (setError(errorText(e)), setLoading(false));
      }
    })();
    return () => {
      cancelled = true;
      clearInterval(refresh);
      void powerbi().then(({ service }) => service.reset(el));
    };
  }, [ref?.id, ref?.groupId, ref?.appId, reloadKey]); // eslint-disable-line react-hooks/exhaustive-deps
  return { element, report, error, loading };
}
