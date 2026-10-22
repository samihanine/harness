import { useEffect, useState } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { BookTextIcon, EyeIcon, LinkIcon, PencilIcon, PlusIcon } from "lucide-react";
import { EntityMenu } from "@/components/entity";
import { Markdown } from "@/components/markdown";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { EntityPicker, Empty, Section, Workspace } from "@/components/workspace";
import { byName, contextsAgent, newContext } from "@/agent/contexts-tools";
import { useLast } from "@/lib/last";
import { patch, put, remove, useCollection, type ContextEntry } from "@/lib/store";

export const Route = createFileRoute("/contexts")({
  validateSearch: (s: Record<string, unknown>): { context?: string; tab?: string } => ({ context: (s.context as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Contexts,
});

/**
 * Contexts: the notes given to the AI with the items they are linked to. Title and text are edited
 * here (by the user or the AI); links are set here or in each item's Context tab.
 */
function Contexts() {
  const { context: contextId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/contexts" });
  useLast("contexts:context", contextId, (context) => navigate({ search: { context }, replace: true }));
  const contexts = useCollection("contexts").sort((a, b) => a.name.localeCompare(b.name));
  const context = contexts.find((c) => c.id === contextId);
  const open = (id: string) => navigate({ search: { context: id } });
  const agent = contextsAgent(() => context, open);
  const create = async () => {
    const name = prompt("Name of the new context")?.trim();
    if (!name) return;
    const existing = byName(contexts, name);
    open(existing?.id ?? (await put("contexts", newContext(name))).id);
  };

  return (
    <Workspace
      title="Contexts"
      icon={BookTextIcon}
      picker={<EntityPicker items={contexts.map((c) => ({ id: c.id, label: c.name }))} value={contextId} onChange={open} placeholder={contexts.length ? "Choose a context" : "New context"} onAdd={() => void create()} addLabel="New context…" />}
      actions={
        <>
          <Button size="sm" variant="outline" onClick={() => void create()}>
            <PlusIcon /> New
          </Button>
          <EntityMenu kind="contexts" entry={context} onRemove={context && (() => confirm(`Delete the context "${context.name}"?`) && void remove("contexts", context.id).then(() => navigate({ search: {} })))} />
        </>
      }
      ai={{ scope: "contexts", agent, placeholder: "Ask to write a context, summarize a document into one, fix a definition…" }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        { id: "edit", label: "Content", icon: PencilIcon, content: context ? <Editor key={context.id} context={context} /> : <List contexts={contexts} onOpen={open} onCreate={() => void create()} /> },
        { id: "preview", label: "Preview", icon: EyeIcon, content: context ? <div className="markdown mx-auto w-full max-w-3xl overflow-y-auto p-8"><Markdown>{context.text || "*Empty*"}</Markdown></div> : <Empty>Choose a context.</Empty> },
        { id: "links", label: "Linked to", icon: LinkIcon, content: context ? <Links context={context} /> : <Empty>Choose a context.</Empty> },
      ]}
    />
  );
}

/** Every context as a card (when none is open). */
function List({ contexts, onOpen, onCreate }: { contexts: ContextEntry[]; onOpen: (id: string) => void; onCreate: () => void }) {
  if (!contexts.length) return <Empty icon={BookTextIcon} action={<Button onClick={onCreate}><PlusIcon /> New context</Button>}>No context yet. A context is a markdown note (rules, definitions) given to the AI with the items it is linked to.</Empty>;
  return (
    <div className="grid content-start gap-3 overflow-y-auto p-6 sm:grid-cols-2 xl:grid-cols-3">
      {contexts.map((c) => (
        <button key={c.id} type="button" onClick={() => onOpen(c.id)} className="flex cursor-pointer flex-col gap-2 rounded-xl border bg-card p-4 text-left shadow-xs transition-shadow hover:shadow-md">
          <span className="font-semibold">{c.name}</span>
          <span className="line-clamp-3 text-[12.5px] text-muted-foreground">{c.text || "Empty"}</span>
          <span className="text-[11.5px] text-muted-foreground">{new Date(c.updatedAt).toLocaleDateString()} · {c.text.length} chars</span>
        </button>
      ))}
    </div>
  );
}

/** Title and markdown text, saved as typed (after a short pause); the AI's changes show at once. */
function Editor({ context }: { context: ContextEntry }) {
  const [name, setName] = useState(context.name);
  const [text, setText] = useState(context.text);
  useEffect(() => setName(context.name), [context.name]);
  useEffect(() => setText(context.text), [context.text]);
  useEffect(() => {
    if (text === context.text && name.trim() === context.name) return;
    const t = setTimeout(() => void patch("contexts", context.id, { text, ...(name.trim() && { name: name.trim() }), updatedAt: Date.now() }), 500);
    return () => clearTimeout(t);
  }, [text, name]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div className="mx-auto flex min-h-0 w-full max-w-4xl flex-1 flex-col gap-3 p-6">
      <Input className="h-11 text-[17px] font-semibold" aria-label="Title" value={name} onChange={(e) => setName(e.target.value)} />
      <Textarea className="min-h-0 flex-1 resize-none font-mono text-[13px] leading-relaxed" aria-label="Markdown text" placeholder="# Rules&#10;- …" value={text} onChange={(e) => setText(e.target.value)} />
      <p className="text-[12px] text-muted-foreground">Markdown · saved automatically · {text.length} chars</p>
    </div>
  );
}

/** The reports, datasets and Excel tables this context is given with. */
function Links({ context }: { context: ContextEntry }) {
  const reports = useCollection("reports");
  const datasets = useCollection("datasets");
  const schemas = useCollection("schemas");
  const toggle = (kind: "reports" | "datasets" | "schemas", item: { id: string; contextIds: string[] }, on: boolean) =>
    void patch(kind, item.id, { contextIds: on ? [...item.contextIds, context.id] : item.contextIds.filter((i) => i !== context.id) });
  const group = (title: string, kind: "reports" | "datasets" | "schemas", items: { id: string; label: string; contextIds: string[] }[]) => (
    <Section title={title}>
      {!items.length && <p className="text-[13px] text-muted-foreground">None in the app.</p>}
      <div className="flex flex-col gap-2">
        {items.map((i) => (
          <label key={i.id} className="flex cursor-pointer items-center gap-3 text-[13.5px]">
            <Switch checked={i.contextIds.includes(context.id)} onCheckedChange={(on) => toggle(kind, i, on)} />
            {i.label}
          </label>
        ))}
      </div>
    </Section>
  );
  return (
    <div className="mx-auto grid w-full max-w-5xl content-start gap-4 overflow-y-auto p-6 md:grid-cols-3">
      {group("Reports", "reports", reports.map((r) => ({ id: r.id, label: r.name, contextIds: r.contextIds })))}
      {group("Datasets", "datasets", datasets.map((d) => ({ id: d.id, label: d.name, contextIds: d.contextIds })))}
      {group("Excel tables", "schemas", schemas.map((s) => ({ id: s.id, label: s.title, contextIds: s.contextIds })))}
    </div>
  );
}
