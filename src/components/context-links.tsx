/**
 * The "Context" tab of a report, a dataset or an Excel file: which contexts (notes given to the AI with
 * it) are linked. A context can be linked to several items; new ones are written in the Contexts screen
 * or right here (title + text).
 */
import { useState } from "react";
import { Link } from "@tanstack/react-router";
import { BookTextIcon, ExternalLinkIcon, PlusIcon } from "lucide-react";
import { Button } from "./ui/button";
import { Input } from "./ui/input";
import { Switch } from "./ui/switch";
import { Textarea } from "./ui/textarea";
import { Empty, Section } from "./workspace";
import { put, useCollection } from "@/lib/store";

export function ContextLinks({ ids, onChange, what }: { ids: string[]; onChange: (ids: string[]) => void; what: string }) {
  const contexts = useCollection("contexts").sort((a, b) => a.name.localeCompare(b.name));
  const [draft, setDraft] = useState<{ name: string; text: string }>();
  const create = async () => {
    if (!draft?.name.trim()) return;
    const c = await put("contexts", { id: crypto.randomUUID(), name: draft.name.trim(), text: draft.text, updatedAt: Date.now() });
    onChange([...ids, c.id]);
    setDraft(undefined);
  };
  return (
    <div className="min-h-0 flex-1 overflow-y-auto p-6">
      <div className="mx-auto flex max-w-3xl flex-col gap-5">
        <Section
          title="Linked contexts"
          description={`Notes given to the AI whenever it works on this ${what}: vocabulary, rules, how the data is used. A context can be linked to several items.`}
          actions={
            <>
              <Button variant="outline" size="sm" asChild>
                <Link to="/contexts" search={{}}>
                  <ExternalLinkIcon /> Contexts
                </Link>
              </Button>
              <Button size="sm" onClick={() => setDraft({ name: "", text: "" })}>
                <PlusIcon /> New context
              </Button>
            </>
          }
        >
          {draft && (
            <div className="mb-4 flex flex-col gap-2 rounded-lg border bg-muted/40 p-3">
              <Input autoFocus placeholder="Title" value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} aria-label="Title of the context" />
              <Textarea placeholder="Content (markdown)" value={draft.text} onChange={(e) => setDraft({ ...draft, text: e.target.value })} className="min-h-28" aria-label="Content of the context" />
              <div className="flex justify-end gap-2">
                <Button variant="ghost" size="sm" onClick={() => setDraft(undefined)}>
                  Cancel
                </Button>
                <Button size="sm" disabled={!draft.name.trim()} onClick={() => void create()}>
                  Create and link
                </Button>
              </div>
            </div>
          )}
          {contexts.length === 0 && !draft ? (
            <Empty icon={BookTextIcon}>No context yet. Create one to give the AI the vocabulary and rules of your business.</Empty>
          ) : (
            <ul className="flex flex-col divide-y">
              {contexts.map((c) => {
                const on = ids.includes(c.id);
                return (
                  <li key={c.id} className="flex items-center gap-3 py-2.5">
                    <BookTextIcon className="size-4 shrink-0 text-muted-foreground" aria-hidden />
                    <div className="min-w-0 flex-1">
                      <div className="truncate font-medium">{c.name}</div>
                      <div className="truncate text-[12.5px] text-muted-foreground">{c.text.replace(/\s+/g, " ").slice(0, 140) || "Empty"}</div>
                    </div>
                    <Switch checked={on} onCheckedChange={(v) => onChange(v ? [...ids, c.id] : ids.filter((x) => x !== c.id))} aria-label={`Link "${c.name}"`} />
                  </li>
                );
              })}
            </ul>
          )}
        </Section>
      </div>
    </div>
  );
}
