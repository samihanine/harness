import { useCallback, useEffect, useRef, useState } from "react";
import { Link, useSearch, useNavigate } from "@tanstack/react-router";
import { EyeIcon, ExternalLinkIcon, FileTextIcon, LinkIcon, PencilIcon, PlusIcon, RefreshCwIcon, SaveIcon } from "lucide-react";
import { diagramSvgText, iconPng, svgToPng } from "@/components/diagram";
import { withToast } from "@/components/entity";
import { Markdown } from "@/components/markdown";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { EntityPicker, Empty, Section, Workspace } from "@/components/workspace";
import { features } from "@/lib/features";
import { useLast } from "@/lib/last";
import { toStorage, type StorageFile } from "@/lib/markdown/storage";
import { patch, useCollection, type Name } from "@/lib/store";
import { attachFile, attachmentNames, createPage, keepMarkdown, listPages, markdownOf, readPage, writePage, type ConfluencePage } from "@/services/confluence";
import { refreshPages } from "@/services/context-pages";
import { errorText } from "@/services/ms";
import { confluenceAgent } from "./agent";
import { confluence } from "./module";

/** Draws and attaches the page's diagrams and icons it does not have yet (the same content, the same file name). */
async function attachMissing(pageId: string, files: StorageFile[]) {
  if (!files.length) return;
  const have = await attachmentNames(pageId);
  for (const f of files.filter((x) => !have.has(x.name))) await attachFile(pageId, f.name, f.kind === "diagram" ? await svgToPng(diagramSvgText(f.code)) : await iconPng(f.icon));
}


/**
 * Confluence: the pages of one folder (Settings › Confluence), read from Confluence each time. The open
 * page is edited as markdown (with mermaid diagrams, panels, statuses, icons: lib/markdown/syntax.ts) —
 * by hand or by the assistant — and written back as the page's next version, its diagrams and icons
 * attached as pictures. A page not changed in Confluence since reads back as the exact markdown it was
 * written from. Each page is also a context: linked to reports, models and Excel tables (in this
 * browser), it is given to the AI with them.
 */
export default function Confluence() {
  const { page: pageId, tab } = useSearch({ from: "/confluence" });
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
      const md = markdownOf(p);
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
    const { storage, files } = toStorage(text);
    await attachMissing(p.id, files);
    const next = await writePage(p, storage);
    keepMarkdown(next, text);
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
      title={confluence.label}
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
            <div className="flex min-h-0 flex-1 gap-4 p-4">
              <div className="flex min-h-0 w-1/2 flex-col gap-2">
                <Textarea className="min-h-0 flex-1 resize-none font-mono text-[12.5px] leading-relaxed" aria-label="Page (markdown)" value={draft} onChange={(e) => setDraft(e.target.value)} />
                <p className="text-[12px] text-muted-foreground">Markdown · {changed ? "unsaved changes" : "saved"} · version {page?.version} · ```mermaid diagrams, &gt; [!NOTE] panels, `status:green:Done`, ![](icon:name)</p>
              </div>
              <div className="min-h-0 w-1/2 overflow-y-auto rounded-lg border bg-card px-6 py-5">
                <Markdown>{draft || "*Empty page*"}</Markdown>
              </div>
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

type Linkable = { id: string; name?: string; title?: string; contextIds: string[] };

/** The items of the features that take contexts (reports, models, Excel tables…) this page is given to the AI with. */
function Links({ pageId }: { pageId: string }) {
  const groups = features().flatMap((f) => (f.linkable ? [f.linkable] : []));
  // A stable list of collections: the same hooks at each render.
  const lists = groups.map((g) => useCollection(g.collection) as unknown as Linkable[]); // eslint-disable-line react-hooks/rules-of-hooks
  const toggle = (kind: Name, item: Linkable, on: boolean) => {
    void patch(kind, item.id, { contextIds: on ? [...item.contextIds, pageId] : item.contextIds.filter((i) => i !== pageId) } as never);
    if (on) void refreshPages([pageId], true);
  };
  return (
    <div className="mx-auto grid w-full max-w-5xl content-start gap-4 overflow-y-auto p-6 md:grid-cols-3">
      {groups.map((g, n) => (
        <Section key={g.collection} title={g.title}>
          {!lists[n].length && <p className="text-[13px] text-muted-foreground">None in the app.</p>}
          <div className="flex flex-col gap-2">
            {lists[n].map((i) => (
              <label key={i.id} className="flex cursor-pointer items-center gap-3 text-[13.5px]">
                <Switch checked={i.contextIds.includes(pageId)} onCheckedChange={(on) => toggle(g.collection, i, on)} />
                {i.name ?? i.title ?? i.id}
              </label>
            ))}
          </div>
        </Section>
      ))}
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
