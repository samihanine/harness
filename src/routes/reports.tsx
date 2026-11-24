/** Reports: a report's structure (downloadable JSON) and its context pages, with an assistant to study it. */
import { useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BarChart3Icon, BookOpenTextIcon, FileJsonIcon } from "lucide-react";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { AddButton, ItemPicker, usePicked } from "@/components/pickers";
import { ContextLinks } from "@/components/context-links";
import { Empty } from "@/components/common";
import { StructureView } from "@/features/views/structure-view";
import { useChat } from "@/lib/chat";
import { contextsOf } from "@/lib/store";

export const Route = createFileRoute("/reports")({ component: Reports });

function Reports() {
  const { items, current, select } = usePicked("report", "reports");
  const [tab, setTab] = useState<"structure" | "context">("structure");
  const itemKey = `report:${current?.id ?? "none"}`;
  const chat = useChat({ screen: "reports", item: current?.id ?? "none", scope: () => ({ report: current, contexts: contextsOf(itemKey) }) });
  return (
    <Workspace
      screen="reports"
      title="Reports"
      item={<ItemPicker kind="report" items={items} current={current} onSelect={select} />}
      tabs={[{ id: "structure", label: "Structure", icon: <FileJsonIcon /> }, { id: "context", label: "Context", icon: <BookOpenTextIcon /> }]}
      tab={tab}
      onTab={setTab}
      chat={<ChatPanel chat={chat} screen="reports" disabled={current ? undefined : "Choose a report first"} suggestions={["Summarise how this report is built", "Which measures does each page use?"]} />}
    >
      {!current ? <Empty icon={<BarChart3Icon />} title="Choose a report"><AddButton kind="report" onAdded={select} /></Empty> : tab === "structure" ? <StructureView report={current} /> : <ContextLinks itemKey={itemKey} title={current.name} />}
    </Workspace>
  );
}
