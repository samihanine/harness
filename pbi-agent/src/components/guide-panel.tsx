import { useEffect, useState } from "react";
import { ChevronRightIcon, FileTextIcon, LinkIcon, PencilIcon, PlusIcon, Trash2Icon, XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { db } from "@/db/db";
import type { Row } from "@/db/schema";
import type { ReportContent } from "@/pbi/types";
import { IconButton } from "./icon-button";
import { Markdown } from "./markdown";
import { SimpleSelect } from "./simple-select";

/** Guide editing (create, edit, link, delete). false = guides are read-only. */
export const GUIDE_EDITING = true;

/** "reportId/pageName/visualName" links of a guide. */
export const linksOf = (row: Row) =>
  String(row.visual_ids ?? "")
    .split(/[;\n]/)
    .map((s) => s.trim())
    .filter(Boolean)
    .map((link) => {
      const [reportId, pageName, visualName] = link.split("/");
      return { link, reportId, pageName, visualName };
    });

export const guidesOfReport = (reportId: string) => db.rows("guides").filter((g) => linksOf(g).some((l) => l.reportId === reportId));

export type Selection = { pageName?: string; visualName?: string };

type View = { kind: "pages" } | { kind: "page"; pageName: string } | { kind: "guide"; id: string };

/**
 * Guides of a report: report > page > guide. A clicked visual opens its guide, or offers to
 * attach it to an existing guide or to create one.
 */
export function GuidePanel({
  reportId,
  reportTitle,
  content,
  selection,
  onGoTo,
}: {
  reportId: string;
  reportTitle: string;
  content?: ReportContent;
  selection: Selection;
  onGoTo: (pageName: string, visualName?: string) => void;
}) {
  const [view, setView] = useState<View>({ kind: "pages" });
  const guides = guidesOfReport(reportId);
  const pages = content?.pages ?? [];
  const pageTitle = (name?: string) => pages.find((p) => p.name === name)?.displayName ?? name ?? "";
  const visualTitle = (pageName?: string, visualName?: string) => {
    const visual = pages.find((p) => p.name === pageName)?.visuals.find((v) => v.name === visualName);
    return visual?.title || visual?.type || visualName || "";
  };

  // Follow the visual clicked in the report.
  useEffect(() => {
    if (!selection.visualName) {
      if (selection.pageName) setView((v) => (v.kind === "guide" ? v : { kind: "page", pageName: selection.pageName! }));
      return;
    }
    const linked = guides.find((g) => linksOf(g).some((l) => l.reportId === reportId && l.visualName === selection.visualName));
    setView(linked ? { kind: "guide", id: linked.id } : { kind: "page", pageName: selection.pageName ?? "" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selection.pageName, selection.visualName]);

  const create = (pageName?: string, visualName?: string) => {
    const [row] = db.tables.guides.insert([{ title: visualName ? `About ${visualTitle(pageName, visualName)}` : "New guide", content: "", visual_ids: [reportId, pageName, visualName].filter(Boolean).join("/") }]);
    setView({ kind: "guide", id: row.id });
  };
  const attach = (guide: Row, link: string) => db.tables.guides.update([{ id: guide.id, values: { visual_ids: [...linksOf(guide).map((l) => l.link), link].join("; ") } }]);

  const crumbs = (
    <div className="flex min-w-0 items-center gap-1 border-b px-3 py-2 text-[12px] text-muted-foreground">
      <button type="button" className="truncate hover:text-foreground" onClick={() => setView({ kind: "pages" })}>
        {reportTitle}
      </button>
      {view.kind !== "pages" && (
        <>
          <ChevronRightIcon className="size-3 shrink-0" />
          {(() => {
            const pageName = view.kind === "page" ? view.pageName : linksOf(db.row("guides", view.id) ?? ({} as Row))[0]?.pageName;
            return (
              <button type="button" className="truncate hover:text-foreground" onClick={() => pageName && setView({ kind: "page", pageName })}>
                {pageTitle(pageName) || "Report"}
              </button>
            );
          })()}
        </>
      )}
      {view.kind === "guide" && (
        <>
          <ChevronRightIcon className="size-3 shrink-0" />
          <span className="truncate text-foreground">{String(db.row("guides", view.id)?.title ?? "")}</span>
        </>
      )}
    </div>
  );

  if (view.kind === "guide") {
    const guide = db.row("guides", view.id);
    if (!guide) return crumbs;
    return (
      <div className="flex flex-col">
        {crumbs}
        <GuideEntry guide={guide} reportId={reportId} pages={pages} selection={selection} visualTitle={visualTitle} pageTitle={pageTitle} onGoTo={onGoTo} onDeleted={() => setView({ kind: "pages" })} />
      </div>
    );
  }

  if (view.kind === "page") {
    const pageGuides = guides.filter((g) => linksOf(g).some((l) => l.reportId === reportId && l.pageName === view.pageName));
    const unmatched = selection.visualName && selection.pageName === view.pageName && !guides.some((g) => linksOf(g).some((l) => l.visualName === selection.visualName));
    return (
      <div className="flex flex-col">
        {crumbs}
        {unmatched && GUIDE_EDITING && (
          <div className="m-3 flex flex-col gap-2 rounded-lg border bg-card p-3 shadow-soft animate-in fade-in-0 slide-in-from-top-1">
            <p className="text-[12px]">
              No guide for <span className="font-medium">{visualTitle(selection.pageName, selection.visualName)}</span> yet.
            </p>
            <Button size="sm" variant="outline" className="w-fit" onClick={() => create(selection.pageName, selection.visualName)}>
              <PlusIcon /> New guide for this visual
            </Button>
            {guides.length > 0 && (
              <SimpleSelect
                size="sm"
                value={null}
                placeholder="Or attach it to a guide…"
                onChange={(id) => {
                  attach(db.row("guides", id)!, `${reportId}/${selection.pageName}/${selection.visualName}`);
                  setView({ kind: "guide", id });
                }}
                options={guides.map((g) => ({ value: g.id, label: String(g.title) }))}
              />
            )}
          </div>
        )}
        <List items={pageGuides.map((g) => ({ id: g.id, label: String(g.title), hint: `${linksOf(g).length} visual(s)`, onClick: () => setView({ kind: "guide", id: g.id }) }))} empty={GUIDE_EDITING ? "No guide on this page: click a visual to write one." : "No guide on this page."} />
        {GUIDE_EDITING && (
          <Button variant="ghost" size="sm" className="m-2 w-fit text-muted-foreground" onClick={() => create(view.pageName)}>
            <PlusIcon /> Page guide
          </Button>
        )}
      </div>
    );
  }

  return (
    <div className="flex flex-col">
      {crumbs}
      <List
        items={pages.map((p) => ({
          id: p.name,
          label: p.displayName,
          hint: `${guides.filter((g) => linksOf(g).some((l) => l.pageName === p.name)).length} guide(s)`,
          onClick: () => {
            setView({ kind: "page", pageName: p.name });
            onGoTo(p.name);
          },
        }))}
        empty="The report content is not read yet: refresh it in the database."
      />
    </div>
  );
}

function List({ items, empty }: { items: { id: string; label: string; hint?: string; onClick: () => void }[]; empty: string }) {
  if (items.length === 0) return <p className="p-4 text-[12px] text-muted-foreground">{empty}</p>;
  return (
    <ul className="p-1.5">
      {items.map((item) => (
        <li key={item.id}>
          <button type="button" onClick={item.onClick} className="flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left hover:bg-muted">
            <FileTextIcon className="size-3.5 shrink-0 text-muted-foreground" />
            <span className="truncate">{item.label}</span>
            {item.hint && <span className="ml-auto shrink-0 text-[11px] text-muted-foreground">{item.hint}</span>}
          </button>
        </li>
      ))}
    </ul>
  );
}

function GuideEntry({
  guide,
  reportId,
  pages,
  selection,
  visualTitle,
  pageTitle,
  onGoTo,
  onDeleted,
}: {
  guide: Row;
  reportId: string;
  pages: ReportContent["pages"];
  selection: Selection;
  visualTitle: (page?: string, visual?: string) => string;
  pageTitle: (page?: string) => string;
  onGoTo: (pageName: string, visualName?: string) => void;
  onDeleted: () => void;
}) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState({ title: String(guide.title ?? ""), content: String(guide.content ?? "") });
  const [confirm, setConfirm] = useState(false);
  useEffect(() => {
    setDraft({ title: String(guide.title ?? ""), content: String(guide.content ?? "") });
    setEditing(GUIDE_EDITING && !guide.content);
    setConfirm(false);
  }, [guide.id]); // eslint-disable-line react-hooks/exhaustive-deps
  const links = linksOf(guide);
  const setLinks = (next: string[]) => db.tables.guides.update([{ id: guide.id, values: { visual_ids: next.join("; ") } }]);
  const save = () => {
    db.tables.guides.update([{ id: guide.id, values: draft }]);
    setEditing(false);
  };
  const candidates = pages.flatMap((p) => p.visuals.map((v) => ({ value: `${reportId}/${p.name}/${v.name}`, label: `${p.displayName} › ${v.title || v.type}` }))).filter((c) => !links.some((l) => l.link === c.value));

  return (
    <div className="flex flex-col gap-3 p-3 animate-in fade-in-0">
      {editing ? (
        <>
          <input value={draft.title} onChange={(e) => setDraft({ ...draft, title: e.target.value })} className="h-8 rounded-md border bg-transparent px-2 font-medium outline-none focus:border-ring" />
          <textarea
            autoFocus
            value={draft.content}
            onChange={(e) => setDraft({ ...draft, content: e.target.value })}
            placeholder="Markdown: what this visual shows, how to read it, caveats…"
            className="field-sizing-content min-h-40 rounded-md border bg-transparent p-2 font-mono text-[12px] outline-none focus:border-ring"
          />
          <div className="flex gap-2">
            <Button size="sm" onClick={save}>
              Save
            </Button>
            <Button size="sm" variant="ghost" onClick={() => setEditing(false)}>
              Cancel
            </Button>
          </div>
        </>
      ) : (
        <>
          <div className="flex items-start gap-2">
            <h3 className="flex-1 text-[14px] font-medium">{String(guide.title)}</h3>
            {GUIDE_EDITING && (
              <IconButton label="Edit" size="icon-xs" onClick={() => setEditing(true)}>
                <PencilIcon />
              </IconButton>
            )}
          </div>
          {guide.content ? <Markdown>{String(guide.content)}</Markdown> : <p className="text-[12px] text-muted-foreground">Empty guide.</p>}
        </>
      )}

      <div className="flex flex-col gap-1.5 border-t pt-3">
        <span className="text-[11px] text-muted-foreground">Linked to</span>
        <div className="flex flex-wrap gap-1">
          {links.map((l) => (
            <span key={l.link} className={cn("inline-flex h-6 items-center gap-1 rounded-md border bg-card pl-2 text-[12px]", !GUIDE_EDITING && "pr-2")}>
              <button type="button" onClick={() => l.pageName && onGoTo(l.pageName, l.visualName)} className="max-w-48 truncate hover:underline">
                {l.reportId !== reportId ? "Other report" : l.visualName ? visualTitle(l.pageName, l.visualName) : pageTitle(l.pageName) || "Report"}
              </button>
              {GUIDE_EDITING && (
                <button type="button" aria-label="Unlink" onClick={() => setLinks(links.filter((x) => x.link !== l.link).map((x) => x.link))} className="px-1 text-muted-foreground hover:text-foreground">
                  <XIcon className="size-3" />
                </button>
              )}
            </span>
          ))}
        </div>
        {GUIDE_EDITING && selection.visualName && !links.some((l) => l.visualName === selection.visualName) && (
          <Button size="xs" variant="outline" className="w-fit" onClick={() => setLinks([...links.map((l) => l.link), `${reportId}/${selection.pageName}/${selection.visualName}`])}>
            <LinkIcon /> Link the selected visual
          </Button>
        )}
        {GUIDE_EDITING && candidates.length > 0 && (
          <SimpleSelect size="sm" value={null} placeholder="Link another visual…" onChange={(value) => setLinks([...links.map((l) => l.link), value])} options={candidates} />
        )}
      </div>

      {GUIDE_EDITING && (
        <div className="border-t pt-3">
          {confirm ? (
            <div className="flex items-center gap-2 text-[12px]">
              Delete this guide?
              <Button
                size="xs"
                variant="destructive"
                onClick={() => {
                  void db.remove("guides", guide.id);
                  onDeleted();
                }}
              >
                Delete
              </Button>
              <Button size="xs" variant="ghost" onClick={() => setConfirm(false)}>
                Cancel
              </Button>
            </div>
          ) : (
            <Button size="xs" variant="ghost" className="text-muted-foreground" onClick={() => setConfirm(true)}>
              <Trash2Icon /> Delete guide
            </Button>
          )}
        </div>
      )}
    </div>
  );
}
