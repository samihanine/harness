/** The Confluence pages linked to an item as its context (as in v1: the links kept in this browser). */
import { useEffect, useState } from "react";
import { BookOpenTextIcon, Loader2Icon, PlusIcon, XIcon } from "lucide-react";
import { Empty, ErrorBox } from "./common";
import { Button } from "./ui/button";
import { errorText, get } from "@/lib/api";
import { useContexts } from "@/lib/store";

export function ContextLinks({ itemKey, title }: { itemKey: string; title?: string }) {
  const [links, setLinks] = useContexts(itemKey);
  const [pages, setPages] = useState<{ id: string; title: string }[]>();
  const [error, setError] = useState<string>();
  const [picking, setPicking] = useState(false);
  const [query, setQuery] = useState("");
  useEffect(() => {
    if (picking && !pages) get("/confluence/pages").then(setPages, (e) => setError(errorText(e)));
  }, [picking, pages]);
  return (
    <div className="mx-auto flex max-w-3xl flex-col gap-4 p-6">
      <div className="flex items-center gap-3">
        <div className="flex-1">
          <div className="text-[15px] font-semibold">Context</div>
          <div className="text-[13px] text-muted-foreground">Confluence pages the assistant reads about {title ?? "this item"}: business rules, definitions, how it is used.</div>
        </div>
        <Button variant="dark" size="sm" onClick={() => setPicking(!picking)}>
          <PlusIcon /> Link a page
        </Button>
      </div>
      {picking && (
        <div className="rounded-2xl border p-3">
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search the folder's pages" className="input mb-2" />
          {error && <ErrorBox>{error}</ErrorBox>}
          {!pages && !error && (
            <div className="flex items-center gap-2 p-2 text-[13px] text-muted-foreground">
              <Loader2Icon className="size-4 animate-spin" /> Loading Confluence pages…
            </div>
          )}
          <div className="max-h-64 overflow-y-auto">
            {(pages ?? [])
              .filter((p) => p.title.toLowerCase().includes(query.toLowerCase()) && !links.some((l) => l.id === p.id))
              .map((p) => (
                <button key={p.id} onClick={() => setLinks([...links, p])} className="flex w-full items-center gap-2 rounded-lg px-2 py-1.5 text-left text-[13.5px] hover:bg-muted">
                  <BookOpenTextIcon className="size-4 text-muted-foreground" /> {p.title}
                </button>
              ))}
          </div>
        </div>
      )}
      {links.length === 0 ? (
        <Empty icon={<BookOpenTextIcon />} title="No context page">Link Confluence pages so the assistant knows the rules and vocabulary of this item.</Empty>
      ) : (
        <div className="flex flex-col gap-2">
          {links.map((l) => (
            <div key={l.id} className="flex items-center gap-3 rounded-2xl border px-4 py-3">
              <BookOpenTextIcon className="size-4 text-muted-foreground" />
              <span className="flex-1 text-[14px]">{l.title}</span>
              <Button variant="ghost" size="icon-sm" onClick={() => setLinks(links.filter((x) => x.id !== l.id))} aria-label="Unlink">
                <XIcon />
              </Button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
