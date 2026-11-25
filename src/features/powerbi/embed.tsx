/**
 * A Power BI report embedded in the page (powerbi-client), with the Power BI token from the server
 * (refreshed before it expires): in view mode, or in the editor (`edit`: the Builder's editor mode, with
 * powerbi-report-authoring). `ready()` resolves once the report has drawn: the assistant's actions wait
 * for it. The report fits the space it has when it opens.
 */
import { useEffect, useRef, useState } from "react";
import type * as PbiClient from "powerbi-client";
import type { Ref } from "@shared/types";
import { errorText, get } from "@/lib/api";

export type Report = PbiClient.Report;
let lib: Promise<{ pbi: typeof PbiClient; service: PbiClient.service.Service }> | undefined;
export const powerbi = () =>
  (lib ??= import("powerbi-client").then(async (pbi) => {
    // The authoring SDK adds the editing methods to the same client.
    await import("powerbi-report-authoring");
    return { pbi, service: new pbi.service.Service(pbi.factories.hpmFactory, pbi.factories.wpmpFactory, pbi.factories.routerFactory) };
  }));

let token: { token: string; expiresAt: number } | undefined;
export async function embedToken() {
  if (!token || token.expiresAt < Date.now() + 120_000) token = await get("/embed-token");
  return token!.token;
}

export const embedUrl = (ref: Ref) => `https://app.powerbi.com/reportEmbed?reportId=${ref.id}${ref.appId ? `&appId=${ref.appId}` : ref.groupId ? `&groupId=${ref.groupId}` : ""}`;

/** The page fitted to the frame: whole when the frame is wider than the page, else as wide as the frame (scrolls). */
export async function fitReport(report: Report, el: HTMLElement | null) {
  const { pbi } = await powerbi();
  const page = (await report.getPages().catch(() => [])).find((p) => p.isActive);
  const size = page?.defaultSize ?? { width: 1280, height: 720 };
  const box = el?.getBoundingClientRect();
  const wide = !box || box.width / Math.max(1, box.height) >= (size.width ?? 16) / (size.height ?? 9);
  await report.setZoom(1).catch(() => undefined);
  await report.updateSettings({ layoutType: pbi.models.LayoutType.Custom, customLayout: { displayOption: wide ? pbi.models.DisplayOption.FitToPage : pbi.models.DisplayOption.FitToWidth } }).catch(() => undefined);
}

export function useEmbed(ref: Ref | undefined, reloadKey = 0, mode: "view" | "edit" = "view") {
  const element = useRef<HTMLDivElement>(null);
  const [report, setReport] = useState<Report>();
  const [error, setError] = useState<string>();
  const [loading, setLoading] = useState(false);
  // Resolved when the report has drawn (re-created at each load).
  const readyRef = useRef<{ promise: Promise<void>; resolve: () => void; reject: (e: Error) => void }>(undefined);
  const newReady = () => {
    let resolve!: () => void, reject!: (e: Error) => void;
    const promise = new Promise<void>((res, rej) => ((resolve = res), (reject = rej)));
    promise.catch(() => undefined);
    readyRef.current = { promise, resolve, reject };
  };
  if (!readyRef.current) newReady();
  useEffect(() => {
    const el = element.current;
    if (!ref || !el) return;
    let cancelled = false;
    newReady();
    setReport(undefined);
    setError(undefined);
    setLoading(true);
    // The server opens the same report in its headless browser: the assistant's first look is then quick.
    void fetch("/api/powerbi/warm", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id: ref.id, groupId: ref.groupId, appId: ref.appId }) }).catch(() => undefined);
    let refresh: ReturnType<typeof setInterval> | undefined;
    void (async () => {
      try {
        const { pbi, service } = await powerbi();
        service.reset(el);
        const edit = mode === "edit";
        const r = service.embed(el, {
          type: "report",
          id: ref.id,
          embedUrl: embedUrl(ref),
          accessToken: await embedToken(),
          tokenType: pbi.models.TokenType.Aad,
          permissions: edit ? pbi.models.Permissions.All : pbi.models.Permissions.Read,
          viewMode: edit ? pbi.models.ViewMode.Edit : pbi.models.ViewMode.View,
          settings: { panes: { filters: { visible: true, expanded: false }, pageNavigation: { visible: true } }, background: pbi.models.BackgroundType.Transparent, layoutType: pbi.models.LayoutType.Custom, customLayout: { displayOption: pbi.models.DisplayOption.FitToPage } },
        }) as Report;
        let first = true;
        r.on("loaded", async () => {
          if (cancelled) return;
          // The first view, to come back to (reset_view).
          (r as Report & { initialState?: string }).initialState = (await r.bookmarksManager.capture().catch(() => undefined))?.state;
          setReport(r);
          void fitReport(r, el);
        });
        r.on("rendered", () => {
          if (cancelled || !first) return;
          first = false;
          setLoading(false);
          readyRef.current?.resolve();
        });
        r.on("error", (e) => {
          const d = e.detail as { message?: string; detailedMessage?: string };
          const message = d?.detailedMessage ?? d?.message ?? "Could not load the report";
          if (!cancelled) (setError(message), setLoading(false), readyRef.current?.reject(new Error(message)));
        });
        refresh = setInterval(async () => {
          if (token && token.expiresAt < Date.now() + 300_000) await r.setAccessToken(await embedToken()).catch(() => undefined);
        }, 60_000);
      } catch (e) {
        if (!cancelled) (setError(errorText(e)), setLoading(false), readyRef.current?.reject(new Error(errorText(e))));
      }
    })();
    return () => {
      cancelled = true;
      clearInterval(refresh);
      void powerbi().then(({ service }) => service.reset(el));
    };
  }, [ref?.id, ref?.groupId, ref?.appId, reloadKey, mode]); // eslint-disable-line react-hooks/exhaustive-deps
  /** Waits for the report to draw (60 s at most). */
  const ready = () =>
    Promise.race([readyRef.current!.promise, new Promise<void>((_, reject) => setTimeout(() => reject(new Error("The report on screen did not finish loading in 60 s.")), 60_000))]);
  return { element, report, error, loading, ready };
}
