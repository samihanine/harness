/** Trainer: a report embedded, and an assistant that reads it, sets its slicers and filters, and explains. */
import { useMemo, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BarChart3Icon, BookOpenTextIcon, DatabaseIcon, ExternalLinkIcon, GraduationCapIcon } from "lucide-react";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { AddButton, ItemPicker, usePicked } from "@/components/pickers";
import { ContextLinks } from "@/components/context-links";
import { Empty } from "@/components/common";
import { Button } from "@/components/ui/button";
import { ReportFrame } from "@/features/views/report-view";
import { ModelView } from "@/features/views/model-view";
import { useEmbed, type Report } from "@/features/powerbi/embed";
import { reportActions } from "@/tools/browser/report-actions";
import { useChat } from "@/lib/chat";
import { contextsOf } from "@/lib/store";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/trainer")({ component: Trainer });

type TabId = "report" | "model" | "context";

function Trainer() {
  const { items, current, select } = usePicked("report", "trainer");
  const [tab, setTab] = useState<TabId>("report");
  const embed = useEmbed(current);
  const reportRef = useRef<Report | undefined>(undefined);
  reportRef.current = embed.report;
  const itemKey = `report:${current?.id ?? "none"}`;
  const chat = useChat({ screen: "trainer", item: current?.id ?? "none", scope: () => ({ report: current, contexts: contextsOf(itemKey) }), actions: () => reportActions(() => reportRef.current) });
  const model = useMemo(() => (current?.datasetId ? { id: current.datasetId, groupId: current.datasetGroupId } : undefined), [current?.datasetId, current?.datasetGroupId]);
  return (
    <Workspace
      screen="trainer"
      title="Trainer"
      item={<ItemPicker kind="report" items={items} current={current} onSelect={select} />}
      actions={current?.webUrl && (
        <Button variant="outline" size="sm" asChild>
          <a href={current.webUrl} target="_blank" rel="noreferrer"><ExternalLinkIcon /> Open in Power BI</a>
        </Button>
      )}
      tabs={[{ id: "report", label: "Report", icon: <BarChart3Icon /> }, { id: "model", label: "Semantic model", icon: <DatabaseIcon /> }, { id: "context", label: "Context", icon: <BookOpenTextIcon /> }]}
      tab={tab}
      onTab={setTab}
      scroll={false}
      chat={<ChatPanel chat={chat} screen="trainer" disabled={current ? undefined : "Choose a report first"} placeholder="Ask about the report, or what to show…" suggestions={["Explain this page", "Show me the figures for my biggest region", "What changed most compared to last year?"]} />}
    >
      {!current ? (
        <Empty icon={<GraduationCapIcon />} title="Choose a report">Add a report you can open in Power BI: the assistant reads it, filters it and explains it with you.<div><AddButton kind="report" onAdded={select} /></div></Empty>
      ) : (
        <>
          {/* The report stays loaded on the other tabs: the assistant keeps acting on it. */}
          <div className={cn("min-h-0 flex-1 flex-col", tab === "report" ? "flex" : "hidden")}>
            <ReportFrame embed={embed} refOf={current} />
          </div>
          {tab === "model" && <div className="min-h-0 flex-1 overflow-auto"><ModelView refOf={model} name={current.name} /></div>}
          {tab === "context" && <div className="min-h-0 flex-1 overflow-auto"><ContextLinks itemKey={itemKey} title={current.name} /></div>}
        </>
      )}
    </Workspace>
  );
}
