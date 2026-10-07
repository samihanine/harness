import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { MonitorPlayIcon } from "lucide-react";
import { z } from "zod";
import { errorResult, textResult, useAgentFeatures, useLatest } from "@/agent/server";
import { GuidePanel, guidesOfReport, linksOf } from "@/components/guide-panel";
import type { Selection } from "@/components/guide-panel";
import { NeedsDb } from "@/components/guards";
import { EmptyState } from "@/components/page";
import { useSideTab } from "@/components/side-panel";
import { SearchSelect } from "@/components/search-select";
import { db, useDb } from "@/db/db";
import type { Row } from "@/db/schema";
import { embedReport, pbi, resetEmbed } from "@/pbi/embed";
import type { ReportContent } from "@/pbi/types";

export const Route = createFileRoute("/viewer")({
  component: () => (
    <NeedsDb>
      <Viewer />
    </NeedsDb>
  ),
});

/** Compact description of the report for the agent. */
const describe = (content?: ReportContent) =>
  content?.pages
    .map(
      (p) =>
        `Page "${p.displayName}" (${p.name}):\n${p.visuals
          .filter((v) => !v.hidden)
          .map((v) => `  - ${v.name}: ${v.type}${v.title ? ` "${v.title}"` : ""} ${Object.entries(v.fields).map(([role, f]) => `${role}=${f.join(", ")}`).join("; ")}`)
          .join("\n")}`,
    )
    .join("\n") ?? "Content not read yet.";

function Viewer() {
  const navigate = useNavigate();
  const data = useDb();
  const reports = data.rows("reports");
  const [reportId, setReportId] = useState<string | null>(() => localStorage.getItem("pbi-agent:viewer"));
  const row = reportId ? data.row("reports", reportId) : undefined;
  const [content, setContent] = useState<ReportContent>();
  const [selection, setSelection] = useState<Selection>({});
  const [error, setError] = useState("");
  const element = useRef<HTMLDivElement>(null);
  const embedded = useRef<pbi.Report | null>(null);
  const latest = useLatest({ row, content, selection });
  const target = useSideTab("guide", "Guide");

  useEffect(() => {
    if (reportId) localStorage.setItem("pbi-agent:viewer", reportId);
  }, [reportId]);
  useEffect(() => {
    void db.readJson<ReportContent>(row?.content_json).then(setContent);
  }, [row?.content_json]);

  // Embed the report and follow pages / clicked visuals.
  useEffect(() => {
    const el = element.current;
    if (!row || !el) return;
    setError("");
    setSelection({});
    let cancelled = false;
    embedReport(el, { reportId: row.id, groupId: (row.workspace_id as string) || undefined })
      .then(async (report) => {
        if (cancelled) return;
        embedded.current = report;
        const page = await report.getActivePage().catch(() => null);
        setSelection({ pageName: page?.name });
        report.on("pageChanged", (e) => setSelection({ pageName: (e.detail as { newPage: { name: string } }).newPage.name }));
        report.on("visualClicked", (e) => {
          const detail = e.detail as { visual: { name: string; page?: { name: string } } };
          setSelection((s) => ({ pageName: detail.visual.page?.name ?? s.pageName, visualName: detail.visual.name }));
        });
      })
      .catch((e: Error) => !cancelled && setError(e.message));
    return () => {
      cancelled = true;
      embedded.current = null;
      resetEmbed(el);
    };
  }, [row?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const goTo = async (pageName: string, visualName?: string) => {
    const report = embedded.current;
    if (!report) return;
    await report.setPage(pageName).catch(() => undefined);
    setSelection({ pageName, visualName });
  };

  useAgentFeatures(
    row
      ? `Report viewer: the user looks at the Power BI report "${String(row.title)}". viewer://report describes its pages and visuals (fields by role), viewer://guides the guides written for it (markdown, linked to visuals "reportId/pageName/visualName"), viewer://selection the current page and clicked visual. You can read a visual's data (get_visual_data), change page (go_to_page) and write guides.`
      : "Report viewer: no report selected.",
    (server) => {
      if (!row) return [];
      const guideJson = (g: Row) => ({ id: g.id, title: g.title, content: g.content, visual_ids: linksOf(g).map((l) => l.link) });
      return [
        server.registerResource("report", "viewer://report", { title: "Report pages and visuals", mimeType: "text/plain", annotations: { priority: 1 } }, async (uri) => ({
          contents: [{ uri: uri.href, text: `${row.context ? `Context: ${String(row.context)}\n` : ""}${describe(latest.current.content)}` }],
        })),
        server.registerResource("guides", "viewer://guides", { title: "Guides of this report", mimeType: "application/jsonl", annotations: { priority: 1 } }, async (uri) => ({
          contents: [{ uri: uri.href, text: guidesOfReport(row.id).map((g) => JSON.stringify(guideJson(g))).join("\n") || "No guide yet." }],
        })),
        server.registerResource("selection", "viewer://selection", { title: "Current page and clicked visual", mimeType: "application/json", annotations: { priority: 1 } }, async (uri) => ({
          contents: [{ uri: uri.href, text: JSON.stringify(latest.current.selection) }],
        })),
        server.registerTool(
          "get_visual_data",
          { description: "Data shown by a visual (summarized, CSV) as the user sees it with the current filters.", inputSchema: z.object({ pageName: z.string(), visualName: z.string() }), annotations: { readOnlyHint: true } },
          async ({ pageName, visualName }) => {
            try {
              const page = (await embedded.current!.getPages()).find((p) => p.name === pageName);
              const visual = (await page?.getVisuals())?.find((v) => v.name === visualName);
              if (!visual) return errorResult(`No visual ${visualName} on page ${pageName}`);
              const { data: csv } = await visual.exportData(pbi.models.ExportDataType.Summarized, 500);
              return textResult(csv);
            } catch (e) {
              return errorResult(e);
            }
          },
        ),
        server.registerTool("go_to_page", { description: "Shows a page of the report.", inputSchema: z.object({ pageName: z.string() }) }, async ({ pageName }) => {
          const before = latest.current.selection.pageName;
          await goTo(pageName);
          return textResult("Page shown.", before ? { name: "go_to_page", arguments: { pageName: before } } : undefined);
        }),
        server.registerTool(
          "save_guide",
          {
            description: "Creates a guide (no id) or updates one (id). visual_ids: links \"reportId/pageName/visualName\" (or \"reportId/pageName\" for a page guide).",
            inputSchema: z.object({ id: z.string().optional(), title: z.string().optional(), content: z.string().optional().describe("markdown"), visual_ids: z.array(z.string()).optional() }),
          },
          async ({ id, title, content: text, visual_ids }) => {
            const values = Object.fromEntries(Object.entries({ title, content: text, visual_ids: visual_ids?.join("; ") }).filter(([, v]) => v !== undefined));
            if (!id) {
              if (!title) return errorResult("A new guide needs a title.");
              const [created] = db.tables.guides.insert([values]);
              return textResult(`Guide ${created.id} created.`, { name: "delete_guide", arguments: { id: created.id } });
            }
            const before = db.row("guides", id);
            if (!before) return errorResult(`No guide ${id}`);
            db.tables.guides.update([{ id, values }]);
            return textResult("Guide updated.", { name: "save_guide", arguments: { id, ...Object.fromEntries(Object.keys(values).map((k) => [k, k === "visual_ids" ? linksOf(before).map((l) => l.link) : before[k]])) } });
          },
        ),
        server.registerTool("delete_guide", { description: "Deletes a guide.", inputSchema: z.object({ id: z.string() }) }, async ({ id }) => {
          const before = db.row("guides", id);
          if (!before) return errorResult(`No guide ${id}`);
          db.tables.guides.remove([id]);
          return textResult("Guide deleted.", { name: "save_guide", arguments: { title: before.title, content: before.content, visual_ids: linksOf(before).map((l) => l.link) } });
        }),
      ];
    },
    [row?.id],
  );

  return (
    <div className="flex h-full flex-col">
      <div className="flex h-11 shrink-0 items-center gap-2 border-b px-3">
        <SearchSelect size="sm" className="w-72" value={reportId} placeholder="Choose a report" onChange={setReportId} onAdd={() => void navigate({ to: "/database" })} addLabel="Add a report" options={reports.map((r) => ({ value: r.id, label: String(r.title), recent: String(r.updated_at ?? "") }))} />
        {error && <span className="truncate text-[12px] text-destructive">{error}</span>}
      </div>
      {row ? (
        <div ref={element} className="min-h-0 flex-1 [&_iframe]:border-0" />
      ) : (
        <div className="mx-auto w-full max-w-lg p-8">
          <EmptyState icon={<MonitorPlayIcon />} title="Choose a report">
            Reports come from the database. Click a visual to read or write its guide.
          </EmptyState>
        </div>
      )}
      {target &&
        createPortal(
          row ? (
            <GuidePanel reportId={row.id} reportTitle={String(row.title)} content={content} selection={selection} onGoTo={(p, v) => void goTo(p, v)} />
          ) : (
            <p className="p-4 text-[12px] text-muted-foreground">Choose a report to see its guides.</p>
          ),
          target,
        )}
    </div>
  );
}
