/** Trainer: a report embedded, and an assistant that reads it, sets its slicers and filters, and explains. */
import { useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BarChart3Icon, DatabaseIcon, ExternalLinkIcon, FileJsonIcon, GraduationCapIcon } from "lucide-react";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { ItemPicker, usePicked } from "@/components/pickers";
import { Empty, ErrorBox, Loading } from "@/components/common";
import { Button } from "@/components/ui/button";
import { ReportFrame } from "@/features/views/report-view";
import { ModelView } from "@/features/views/model-view";
import { StructureView } from "@/features/views/structure-view";
import { captureView, showView, useEmbed, type Report } from "@/features/powerbi/embed";
import { reportActions } from "@/tools/browser/report-actions";
import { useChat } from "@/lib/chat";
import { screenSearch } from "@/lib/params";
import { withStructure } from "@/lib/items";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/trainer")({ component: Trainer, validateSearch: screenSearch });

type TabId = "report" | "structure" | "model";

function Trainer() {
  const { items, entry, current, select, loading, error } = usePicked("trainer");
  const [tab, setTab] = useState<TabId>("report");
  const embed = useEmbed(current);
  const reportRef = useRef<Report | undefined>(undefined);
  reportRef.current = embed.report;
  const chat = useChat({
    screen: "trainer",
    item: current?.id ?? "none",
    scope: async () => ({ report: await withStructure(current) }),
    actions: () => reportActions(() => reportRef.current, embed.ready),
    view: () => captureView(reportRef.current),
    restoreView: (state) => showView(() => reportRef.current, current!.id, embed.ready, state),
  });
  const model = useMemo(() => (current?.datasetId ? { id: current.datasetId, groupId: current.datasetGroupId } : undefined), [current?.datasetId, current?.datasetGroupId]);
  return (
    <Workspace
      screen="trainer"
      title="Trainer"
      item={<ItemPicker kind="report" current={entry} onSelect={select} />}
      actions={current?.webUrl && (
        <Button variant="outline" size="sm" asChild>
          <a href={current.webUrl} target="_blank" rel="noreferrer"><ExternalLinkIcon /> Power BI</a>
        </Button>
      )}
      tabs={[{ id: "report", label: "Report", icon: <BarChart3Icon /> }, { id: "structure", label: "Structure", icon: <FileJsonIcon /> }, { id: "model", label: "Semantic model", icon: <DatabaseIcon /> }]}
      tab={tab}
      onTab={setTab}
      scroll={false}
      chat={<ChatPanel chat={chat} screen="trainer" itemLabel={(id) => items.find((r) => r.id === id)?.name} disabled={current ? undefined : "Choose a report first"} placeholder="Ask about the report, or what to show…" suggestions={["Explain this page", "Show me the figures for my biggest region", "What changed most compared to last year?"]} />}
    >
      {error ? (
        <div className="p-6"><ErrorBox>{error}</ErrorBox></div>
      ) : loading ? (
        <Loading label={entry ? `Opening "${entry.name}"…` : "Reading the catalog…"} />
      ) : !current ? (
        <Empty icon={<GraduationCapIcon />} title="Choose a report">Pick a report of the team's catalog: the assistant reads it, filters it and explains it with you.</Empty>
      ) : (
        <>
          {/* The report stays loaded on the other tabs: the assistant keeps acting on it. */}
          <div className={cn("min-h-0 flex-1 flex-col", tab === "report" ? "flex" : "hidden")}>
            <ReportFrame embed={embed} refOf={current} />
          </div>
          {tab === "structure" && <div className="min-h-0 flex-1 overflow-auto"><StructureView report={current} /></div>}
          {tab === "model" && <div className="min-h-0 flex-1 overflow-auto"><ModelView refOf={model} name={current.name} /></div>}
        </>
      )}
    </Workspace>
  );
}
