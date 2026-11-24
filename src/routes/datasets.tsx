/** Semantic models: structure (downloadable JSON with value examples), a DAX console, context pages, and an assistant. */
import { useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BookOpenTextIcon, CodeIcon, DatabaseIcon, Loader2Icon, PlayIcon, TableIcon } from "lucide-react";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { AddButton, ItemPicker, usePicked } from "@/components/pickers";
import { ContextLinks } from "@/components/context-links";
import { Empty, ErrorBox } from "@/components/common";
import { Button } from "@/components/ui/button";
import { ModelView } from "@/features/views/model-view";
import { useChat } from "@/lib/chat";
import { contextsOf, useLocal } from "@/lib/store";
import { errorText, post } from "@/lib/api";

export const Route = createFileRoute("/datasets")({ component: Datasets });

function DaxConsole({ id, groupId }: { id: string; groupId?: string }) {
  const [query, setQuery] = useLocal(`dax:${id}`, "EVALUATE\nTOPN(10, INFO.VIEW.TABLES())");
  const [rows, setRows] = useState<Record<string, unknown>[]>();
  const [error, setError] = useState<string>();
  const [busy, setBusy] = useState(false);
  const run = async () => {
    setBusy(true);
    setError(undefined);
    try {
      setRows(await post("/dax", { id, groupId, query }));
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  const cols = rows?.length ? Object.keys(rows[0]) : [];
  return (
    <div className="flex flex-col gap-3 p-6">
      <textarea value={query} onChange={(e) => setQuery(e.target.value)} spellCheck={false} className="min-h-40 rounded-2xl border bg-muted/30 p-4 font-mono text-[12.5px] outline-none focus:border-foreground/30" onKeyDown={(e) => (e.ctrlKey || e.metaKey) && e.key === "Enter" && void run()} />
      <Button variant="dark" size="sm" className="self-start" onClick={run} disabled={busy}>
        {busy ? <Loader2Icon className="animate-spin" /> : <PlayIcon />} Run (Ctrl+Enter)
      </Button>
      {error && <ErrorBox>{error}</ErrorBox>}
      {rows && (
        <div className="overflow-auto rounded-2xl border">
          <table className="w-full text-[12.5px]">
            <thead className="bg-muted/50"><tr>{cols.map((c) => <th key={c} className="th">{c}</th>)}</tr></thead>
            <tbody>{rows.map((r, i) => <tr key={i} className="border-t">{cols.map((c) => <td key={c} className="px-3 py-1.5">{String(r[c] ?? "")}</td>)}</tr>)}</tbody>
          </table>
          {!rows.length && <div className="p-3 text-[13px] text-muted-foreground">No rows.</div>}
        </div>
      )}
    </div>
  );
}

function Datasets() {
  const { items, current, select } = usePicked("dataset", "datasets");
  const [tab, setTab] = useState<"schema" | "dax" | "context">("schema");
  const [reload, setReload] = useState(0);
  const itemKey = `dataset:${current?.id ?? "none"}`;
  const chat = useChat({ screen: "datasets", item: current?.id ?? "none", scope: () => ({ dataset: current, contexts: contextsOf(itemKey) }), onEffect: (n) => n === "model" && setReload((x) => x + 1) });
  return (
    <Workspace
      screen="datasets"
      title="Semantic models"
      item={<ItemPicker kind="dataset" items={items} current={current} onSelect={select} />}
      tabs={[{ id: "schema", label: "Structure", icon: <TableIcon /> }, { id: "dax", label: "DAX", icon: <CodeIcon /> }, { id: "context", label: "Context", icon: <BookOpenTextIcon /> }]}
      tab={tab}
      onTab={setTab}
      chat={<ChatPanel chat={chat} screen="datasets" disabled={current ? undefined : "Choose a semantic model first"} suggestions={["Explain how the main measures are computed", "Add a year-over-year growth measure", "Which tables are not related to anything?"]} />}
    >
      {!current ? <Empty icon={<DatabaseIcon />} title="Choose a semantic model"><AddButton kind="dataset" onAdded={select} /></Empty> : tab === "schema" ? <ModelView key={reload} refOf={current} name={current.name} /> : tab === "dax" ? <DaxConsole id={current.id} groupId={current.groupId} /> : <ContextLinks itemKey={itemKey} title={current.name} />}
    </Workspace>
  );
}
