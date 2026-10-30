import { useCallback, useEffect, useRef, useState } from "react";
import { Link, createFileRoute, useNavigate } from "@tanstack/react-router";
import { EyeIcon, ExternalLinkIcon, FileTextIcon, LinkIcon, PencilIcon, PlusIcon, RefreshCwIcon, SaveIcon } from "lucide-react";
import { withToast } from "@/components/entity";
import { Markdown } from "@/components/markdown";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { EntityPicker, Empty, Section, Workspace } from "@/components/workspace";
import { SCREEN_LABELS } from "@/config";
import { confluenceAgent } from "@/features/confluence/agent";
import { refreshPages } from "@/features/confluence/context-pages";
import { toMarkdown, toStorage } from "@/features/confluence/markdown";
import { useLast } from "@/lib/last";
import { patch, useCollection } from "@/lib/store";
import { errorText } from "@/services/ms";
import { createPage, listPages, readPage, writePage, type ConfluencePage } from "@/services/confluence";

export const Route = createFileRoute("/confluence")({
  validateSearch: (s: Record<string, unknown>): { page?: string; tab?: string } => ({ page: (s.page as string) || undefined, tab: (s.tab as string) || undefined }),
  component: Confluence,
});

/**
 * Confluence: the pages of one folder (Settings › Confluence), read from Confluence each time. The open
 * page is edited as markdown — by hand or by the assistant — and written back as the page's next
 * version. Each page is also a context: linked to reports, models and Excel tables (in this browser),
 * it is given to the AI with them.
 */
function Confluence() {
  const { page: pageId, tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/confluence" });
  useLast("confluence:page", pageId, (page) => navigate({ search: { page }, replace: true }));
  const [pages, setPages] = useState<{ id: string; title: string }[]>([]);
  const [listError, setListError] = useState<string>();
  const [page, setPage] = useState<ConfluencePage>();
  const [draft, setDraft] = useState("");
  const [saved, setSaved] = useState("");
  const [error, setError] = useState<string>();
  // Read by the agent's tools at call time.
  const live = useRef({ page, draft, saved, pages });
  live.current = { page, draft, saved, pages };

  const loadList = useCallback(() => listPages().then((p) => (setPages(p), setListError(undefined)), (e) => setListError(errorText(e))), []);
  useEffect(() => void loadList(), [loadList]);
  const loadPage = useCallback(async () => {
    setPage(undefined);
    setError(undefined);
    if (!pageId) return;
    try {
      const p = await readPage(pageId);
      const md = toMarkdown(p.html);
      setPage(p);
      setDraft(md);
      setSaved(md);
    } catch (e) {
      setError(errorText(e));
    }
  }, [pageId]);
  useEffect(() => void loadPage(), [loadPage]);

  /** Writes the draft as the page's next version. */
  const save = async () => {
    const { page: p, draft: text } = live.current;
    if (!p) throw new Error("No page is open.");
    const next = await writePage(p, toStorage(text));
    setPage(next);
    setSaved(text);
    void refreshPages([next.id], true);
    return `Page "${next.title}" saved in Confluence (version ${next.version}).`;
  };
  const create = async () => {
    const title = prompt("Title of the new page")?.trim();
    if (!title) return;
    const p = await createPage(title);
    await loadList();
    navigate({ search: { page: p.id } });
  };
  const agent = confluenceAgent(() => ({ title: live.current.page?.title, pages: live.current.pages, draft: () => live.current.draft, saved: () => live.current.saved, setDraft: (t) => ((live.current.draft = t), setDraft(t)), save }));
  const changed = !!page && draft !== saved;

  return (
    <Workspace
      title={SCREEN_LABELS.confluence}
      icon={FileTextIcon}
      picker={<EntityPicker items={pages.map((p) => ({ id: p.id, label: p.title }))} value={pageId} onChange={(page) => navigate({ search: { page } })} placeholder={pages.length ? "Choose a page" : "No page"} onAdd={() => void withToast("Creating the page…", create(), "Page created.")} addLabel="New page…" />}
      actions={
        <>
          {changed && (
            <Button size="sm" variant="ghost" onClick={() => setDraft(saved)}>
              Discard
            </Button>
          )}
          <Button size="sm" disabled={!changed} onClick={() => void withToast("Saving…", save(), "Saved.")}>
            <SaveIcon /> Save
          </Button>
          <Button variant="ghost" size="icon" className="size-9" aria-label="Read again from Confluence" title="Read again from Confluence" onClick={() => void (loadList(), loadPage())}>
            <RefreshCwIcon />
          </Button>
          {page?.webUrl && (
            <Button variant="outline" size="sm" asChild>
              <a href={page.webUrl} target="_blank" rel="noreferrer">
                <ExternalLinkIcon /> Confluence
              </a>
            </Button>
          )}
          <Button size="sm" variant="outline" onClick={() => void withToast("Creating the page…", create(), "Page created.")}>
            <PlusIcon /> New
          </Button>
        </>
      }
      ai={{ scope: `confluence:${pageId}`, agent, placeholder: "Ask to rewrite a section, add a table, fix the tone…", empty: "Choose a page to edit it with the assistant." }}
      tab={tab}
      onTab={(t) => navigate({ search: (s) => ({ ...s, tab: t }), replace: true })}
      tabs={[
        {
          id: "edit",
          label: "Edit",
          icon: PencilIcon,
          content: <Body error={listError ?? error} ready={!!page} chosen={!!pageId}>
            <div className="flex min-h-0 flex-1 flex-col gap-2 p-6">
              <Textarea className="min-h-0 flex-1 resize-none font-mono text-[13px] leading-relaxed" aria-label="Page (markdown)" value={draft} onChange={(e) => setDraft(e.target.value)} />
              <p className="text-[12px] text-muted-foreground">Markdown · {changed ? "unsaved changes" : "saved"} · version {page?.version}</p>
            </div>
          </Body>,
        },
        {
          id: "preview",
          label: "Preview",
          icon: EyeIcon,
          content: <Body error={listError ?? error} ready={!!page} chosen={!!pageId}>
            <div className="mx-auto w-full max-w-3xl overflow-y-auto p-8">
              <Markdown>{draft || "*Empty page*"}</Markdown>
            </div>
          </Body>,
        },
        {
          id: "links",
          label: "Linked to",
          icon: LinkIcon,
          content: <Body error={listError} ready={!!pageId} chosen={!!pageId}>{pageId && <Links pageId={pageId} />}</Body>,
        },
      ]}
    />
  );
}

/** The reports, datasets and Excel tables this page is given to the AI with (as their context). */
function Links({ pageId }: { pageId: string }) {
  const reports = useCollection("reports");
  const datasets = useCollection("datasets");
  const schemas = useCollection("schemas");
  const toggle = (kind: "reports" | "datasets" | "schemas", item: { id: string; contextIds: string[] }, on: boolean) => {
    void patch(kind, item.id, { contextIds: on ? [...item.contextIds, pageId] : item.contextIds.filter((i) => i !== pageId) });
    if (on) void refreshPages([pageId], true);
  };
  const group = (title: string, kind: "reports" | "datasets" | "schemas", items: { id: string; label: string; contextIds: string[] }[]) => (
    <Section title={title}>
      {!items.length && <p className="text-[13px] text-muted-foreground">None in the app.</p>}
      <div className="flex flex-col gap-2">
        {items.map((i) => (
          <label key={i.id} className="flex cursor-pointer items-center gap-3 text-[13.5px]">
            <Switch checked={i.contextIds.includes(pageId)} onCheckedChange={(on) => toggle(kind, i, on)} />
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

/** The tab's content once the page is read; else why not (no setup, an error, nothing chosen, loading). */
function Body({ error, ready, chosen, children }: { error?: string; ready: boolean; chosen: boolean; children: React.ReactNode }) {
  if (error)
    return (
      <Empty icon={FileTextIcon} action={<Button variant="outline" asChild><Link to="/settings" search={{ tab: "confluence" }}>Settings › Confluence</Link></Button>}>
        {error}
      </Empty>
    );
  if (!chosen) return <Empty icon={FileTextIcon}>Choose a page of the folder, or create one.</Empty>;
  if (!ready) return <Empty>Reading the page…</Empty>;
  return <>{children}</>;
}
