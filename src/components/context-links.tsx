/**
 * The "Context" tab of a report, a dataset or an Excel file: which Confluence pages (of the folder set
 * in Settings › Confluence) are given to the AI with it. Links are kept in this browser; each linked
 * page's content too (src/features/confluence/context-pages.ts), read again at each conversation.
 */
import { useCallback, useEffect, useState } from "react";
import { Link } from "@tanstack/react-router";
import { ExternalLinkIcon, FileTextIcon, RefreshCwIcon } from "lucide-react";
import { Button } from "./ui/button";
import { Switch } from "./ui/switch";
import { Empty, Section } from "./workspace";
import { withToast } from "./entity";
import { refreshPages, useContextPages } from "@/features/confluence/context-pages";
import { listPages } from "@/services/confluence";
import { errorText } from "@/services/ms";

const ago = (at: number) => (at ? new Date(at).toLocaleString(undefined, { dateStyle: "short", timeStyle: "short" }) : "never");

export function ContextLinks({ ids, onChange, what }: { ids: string[]; onChange: (ids: string[]) => void; what: string }) {
  const kept = useContextPages();
  const [pages, setPages] = useState<{ id: string; title: string }[]>();
  const [error, setError] = useState<string>();
  const load = useCallback(() => listPages().then((p) => (setPages(p), setError(undefined)), (e) => setError(errorText(e))), []);
  useEffect(() => void load(), [load]);
  // The folder's pages, and the linked pages no longer in it (moved, or Confluence unreachable): still shown.
  const listed = [...(pages ?? []), ...ids.filter((id) => !pages?.some((p) => p.id === id)).map((id) => ({ id, title: kept[id]?.title ?? `Page ${id}` }))];
  const toggle = (id: string, on: boolean) => {
    onChange(on ? [...ids, id] : ids.filter((x) => x !== id));
    if (on) void refreshPages([id], true);
  };
  return (
    <div className="min-h-0 flex-1 overflow-y-auto p-6">
      <div className="mx-auto flex max-w-3xl flex-col gap-5">
        <Section
          title="Linked Confluence pages"
          description={`Pages given to the AI whenever it works on this ${what}: vocabulary, rules, how the data is used. Read again from Confluence at each new conversation; the copy kept in this browser is used when Confluence cannot be reached.`}
          actions={
            <>
              <Button variant="ghost" size="icon" className="size-8" aria-label="Read the pages again" title="Read the pages again" onClick={() => void withToast("Reading the pages…", Promise.all([load(), refreshPages(ids, true)]), "Pages read again.")}>
                <RefreshCwIcon />
              </Button>
              <Button variant="outline" size="sm" asChild>
                <Link to="/confluence" search={{}}>
                  <ExternalLinkIcon /> Confluence
                </Link>
              </Button>
            </>
          }
        >
          {error && (
            <Empty icon={FileTextIcon} action={<Button variant="outline" size="sm" asChild><Link to="/settings" search={{ tab: "confluence" }}>Settings › Confluence</Link></Button>}>
              {error}
            </Empty>
          )}
          {!error && !pages && <Empty>Reading the pages of the folder…</Empty>}
          {pages?.length === 0 && !ids.length && <Empty icon={FileTextIcon}>No page in the Confluence folder yet.</Empty>}
          {listed.length > 0 && (
            <ul className="flex flex-col divide-y">
              {listed.map((p) => {
                const on = ids.includes(p.id);
                const copy = kept[p.id];
                return (
                  <li key={p.id} className="flex items-center gap-3 py-2.5">
                    <FileTextIcon className="size-4 shrink-0 text-muted-foreground" aria-hidden />
                    <div className="min-w-0 flex-1">
                      <div className="truncate font-medium">{p.title}</div>
                      {on && (
                        <div className={`truncate text-[12.5px] ${copy?.error ? "text-destructive" : "text-muted-foreground"}`}>
                          {copy?.error ? `Not read (${copy.error}); copy of ${ago(copy.at)} used` : copy ? `Read ${ago(copy.at)} · ${copy.text.replace(/\s+/g, " ").slice(0, 120) || "empty"}` : "Not read yet"}
                        </div>
                      )}
                    </div>
                    <Switch checked={on} onCheckedChange={(v) => toggle(p.id, v)} aria-label={`Link "${p.title}"`} />
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
