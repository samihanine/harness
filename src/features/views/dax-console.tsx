/** A DAX console on a semantic model (the app's limits apply: rows, time). */
import { useState } from "react";
import { Loader2Icon, PlayIcon } from "lucide-react";
import { Empty, ErrorBox } from "@/components/common";
import { Button } from "@/components/ui/button";
import { errorText, post } from "@/lib/api";
import { useLocal } from "@/lib/store";
import { cn } from "@/lib/utils";

export function DaxConsole({ models }: { models: { id: string; groupId?: string; name: string }[] }) {
  const [picked, setPicked] = useLocal("dax:model", "");
  const model = models.find((m) => m.id === picked) ?? models[0];
  const [query, setQuery] = useLocal(`dax:${model?.id}`, "EVALUATE\nTOPN(10, INFO.VIEW.TABLES())");
  const [rows, setRows] = useState<Record<string, unknown>[]>();
  const [error, setError] = useState<string>();
  const [busy, setBusy] = useState(false);
  if (!model) return <Empty title="No semantic model">Add one in the Sources tab.</Empty>;
  const run = async () => {
    setBusy(true);
    setError(undefined);
    try {
      setRows(await post("/dax", { id: model.id, groupId: model.groupId, query }));
    } catch (e) {
      setError(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  const cols = rows?.length ? Object.keys(rows[0]) : [];
  return (
    <div className="flex flex-col gap-3 p-6">
      {models.length > 1 && (
        <div className="flex flex-wrap gap-1.5">
          {models.map((m) => (
            <button key={m.id} onClick={() => setPicked(m.id)} className={cn("rounded-full border px-3 py-1 text-[12.5px]", m.id === model.id ? "border-foreground bg-foreground text-background" : "hover:bg-muted")}>{m.name}</button>
          ))}
        </div>
      )}
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
