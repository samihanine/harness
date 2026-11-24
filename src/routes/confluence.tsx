/** Confluence: the documentation pages of the folder, edited as markdown with the assistant, published by the user. */
import { useEffect, useMemo, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BookOpenTextIcon, CheckIcon, ChevronDownIcon, EyeIcon, Link2Icon, Loader2Icon, PencilLineIcon, PlusIcon, SendIcon } from "lucide-react";
import { toast } from "sonner";
import { Workspace } from "@/components/workspace";
import { ChatPanel } from "@/components/chat-panel";
import { Empty, ErrorBox, Loading } from "@/components/common";
import { Markdown } from "@/components/markdown";
import { Button } from "@/components/ui/button";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { useChat } from "@/lib/chat";
import { errorText, get, post, put } from "@/lib/api";
import { useItems, useLocal } from "@/lib/store";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/confluence")({ component: Confluence });

type Draft = { pageId: string; title: string; markdown: string; version: number };

function LinkedTo({ pageId }: { pageId: string }) {
  const reports = useItems("report") ?? [];
  const datasets = useItems("dataset") ?? [];
  const decks = useItems("deck") ?? [];
  const analyses = useItems("analysis") ?? [];
  const names: Record<string, string> = Object.fromEntries([...reports.map((r) => [`report:${r.id}`, `Report — ${r.name}`]), ...datasets.map((d) => [`dataset:${d.id}`, `Model — ${d.name}`]), ...decks.map((d) => [`deck:${d.id}`, `Deck — ${d.title}`]), ...analyses.map((a) => [`analysis:${a.id}`, `Analysis — ${a.title}`])]);
  const linked = Object.keys(localStorage)
    .filter((k) => k.startsWith("hub:contexts:"))
    .filter((k) => (JSON.parse(localStorage.getItem(k) ?? "[]") as { id: string }[]).some((l) => l.id === pageId))
    .map((k) => k.slice("hub:contexts:".length));
  return (
    <div className="mx-auto flex max-w-3xl flex-col gap-2 p-6">
      <div className="mb-2 text-[13px] text-muted-foreground">Items using this page as context (link it from their Context tab).</div>
      {linked.length ? linked.map((k) => <div key={k} className="rounded-2xl border px-4 py-3 text-[14px]">{names[k] ?? k}</div>) : <Empty icon={<Link2Icon />} title="Not linked yet" />}
    </div>
  );
}

function Confluence() {
  const [pages, setPages] = useState<{ id: string; title: string }[]>();
  const [error, setError] = useState<string>();
  const [pageId, setPageId] = useLocal("confluence:page", "");
  const [draft, setDraft] = useLocal<Draft | null>(`confluence:draft:${pageId}`, null);
  const [tab, setTab] = useState<"edit" | "preview" | "linked">("edit");
  const [busy, setBusy] = useState<string>();
  const [loadError, setLoadError] = useState<string>();
  useEffect(() => void get("/confluence/pages").then(setPages, (e) => setError(errorText(e))), []);
  // The page read from Confluence when there is no draft of it.
  useEffect(() => {
    if (!pageId || draft) return;
    setLoadError(undefined);
    get<{ id: string; title: string; markdown: string; version: number }>(`/confluence/page/${pageId}`).then((p) => setDraft({ pageId: p.id, title: p.title, markdown: p.markdown, version: p.version }), (e) => setLoadError(errorText(e)));
  }, [pageId, draft]); // eslint-disable-line react-hooks/exhaustive-deps
  const chat = useChat({ screen: "confluence", item: pageId || "none", scope: () => ({ pageId, doc: draft ?? undefined }), onEffect: (n, d) => n === "doc" && (setDraft(d as Draft), setTab("preview")) });
  const current = pages?.find((p) => p.id === pageId);
  const publish = async () => {
    if (!draft) return;
    setBusy("publish");
    try {
      const p = await put<{ version: number }>(`/confluence/page/${draft.pageId}`, draft);
      setDraft({ ...draft, version: p.version });
      toast.success("Published to Confluence");
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(undefined);
    }
  };
  const create = async () => {
    const title = prompt("Title of the new page");
    if (!title) return;
    try {
      const p = await post<{ id: string; title: string }>("/confluence/pages", { title });
      setPages([...(pages ?? []), { id: p.id, title: p.title }]);
      setPageId(p.id);
    } catch (e) {
      toast.error(errorText(e));
    }
  };
  const picker = (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <button className="flex max-w-[420px] items-center gap-1.5 rounded-full px-2 py-1 text-[26px] font-light tracking-tight hover:bg-muted">
          <span className="truncate">{current?.title ?? draft?.title ?? "Choose a page"}</span>
          <ChevronDownIcon className="size-5 text-muted-foreground" />
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="max-h-[60vh] w-80 overflow-y-auto">
        {pages?.map((p) => (
          <DropdownMenuItem key={p.id} onSelect={() => setPageId(p.id)}>
            <CheckIcon className={cn("size-4", p.id === pageId ? "opacity-100" : "opacity-0")} /> {p.title}
          </DropdownMenuItem>
        ))}
        <DropdownMenuSeparator />
        <DropdownMenuItem onSelect={create}><PlusIcon /> New page</DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
  );
  const dirty = useMemo(() => !!draft, [draft]);
  return (
    <Workspace
      screen="confluence"
      title="Confluence"
      item={picker}
      actions={draft && (
        <Button variant="dark" size="sm" onClick={publish} disabled={!!busy || !dirty}>
          {busy ? <Loader2Icon className="animate-spin" /> : <SendIcon />} Publish
        </Button>
      )}
      tabs={[{ id: "edit", label: "Edit", icon: <PencilLineIcon /> }, { id: "preview", label: "Preview", icon: <EyeIcon /> }, { id: "linked", label: "Linked to", icon: <Link2Icon /> }]}
      tab={tab}
      onTab={setTab}
      scroll={tab !== "edit"}
      chat={<ChatPanel chat={chat} screen="confluence" disabled={draft ? undefined : "Choose a page first"} suggestions={["Rewrite this page more clearly", "Document the measures of the model", "Add a FAQ section"]} />}
    >
      {error ? (
        <div className="p-6"><ErrorBox>{error}</ErrorBox></div>
      ) : !pageId ? (
        <Empty icon={<BookOpenTextIcon />} title="Choose a page">The pages of the Confluence folder set in Settings › Confluence.</Empty>
      ) : loadError ? (
        <div className="p-6"><ErrorBox>{loadError}</ErrorBox></div>
      ) : !draft ? (
        <Loading />
      ) : tab === "edit" ? (
        <div className="flex h-full flex-col gap-3 p-5">
          <input value={draft.title} onChange={(e) => setDraft({ ...draft, title: e.target.value })} className="bg-transparent text-[20px] font-semibold outline-none" />
          <textarea value={draft.markdown} onChange={(e) => setDraft({ ...draft, markdown: e.target.value })} spellCheck={false} className="min-h-0 flex-1 rounded-2xl border bg-muted/30 p-4 font-mono text-[12.5px] leading-relaxed outline-none focus:border-foreground/30" />
        </div>
      ) : tab === "preview" ? (
        <div className="mx-auto max-w-3xl p-8">
          <h1 className="mb-4 text-[24px] font-semibold">{draft.title}</h1>
          <Markdown>{draft.markdown}</Markdown>
        </div>
      ) : (
        <LinkedTo pageId={pageId} />
      )}
    </Workspace>
  );
}
