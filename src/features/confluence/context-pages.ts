/**
 * Contexts are Confluence pages: an item (report, model, Excel table) links page ids (`contextIds`),
 * kept in this browser. Each page's content is kept in localStorage as markdown, read again from
 * Confluence once per conversation; when Confluence cannot be reached, the copy kept is used.
 */
import { useSyncExternalStore } from "react";
import { readPage } from "@/services/confluence";
import { toMarkdown } from "./markdown";

export type ContextPage = { id: string; title: string; text: string; at: number; error?: string };

const KEY = "hub-context-pages";
/** One page kept up to this length (localStorage holds a few MB in all). */
const MAX_PAGE = 60_000;

let pages: Record<string, ContextPage> | undefined;
const listeners = new Set<() => void>();
const kept = () => {
  if (!pages) {
    try {
      pages = JSON.parse(localStorage.getItem(KEY) ?? "{}") ?? {};
    } catch {
      pages = {};
    }
  }
  return pages!;
};
const keep = (next: Record<string, ContextPage>) => {
  pages = next;
  try {
    localStorage.setItem(KEY, JSON.stringify(next));
  } catch {}
  listeners.forEach((l) => l());
};

/** The pages kept in this browser (re-rendered when one is read again). */
export const useContextPages = () => useSyncExternalStore((l) => (listeners.add(l), () => void listeners.delete(l)), kept);

/** Pages read from Confluence in this conversation: not read again before the next one. */
const fresh = new Set<string>();
let round: string | undefined;
/** Called at each message (src/agent/loop.ts): in another conversation, the linked pages are read again. */
export const contextRound = (conversationId: string) => void (round !== conversationId && ((round = conversationId), fresh.clear()));

/** Reads the pages from Confluence (those not read in this conversation); a failure keeps the copy, with the error. */
export async function refreshPages(ids: string[], force = false) {
  const todo = ids.filter((id) => force || !fresh.has(id));
  if (!todo.length) return;
  const read = await Promise.all(
    todo.map(async (id): Promise<ContextPage> => {
      fresh.add(id); // tried once per conversation, even when Confluence fails
      try {
        const p = await readPage(id);
        return { id, title: p.title, text: toMarkdown(p.html).slice(0, MAX_PAGE), at: Date.now() };
      } catch (e) {
        const old = kept()[id];
        return { id, title: old?.title ?? `Page ${id}`, text: old?.text ?? "", at: old?.at ?? 0, error: e instanceof Error ? e.message : String(e) };
      }
    }),
  );
  keep({ ...kept(), ...Object.fromEntries(read.map((p) => [p.id, p])) });
}

/** A page's title for the lists (the copy kept; its id when never read). */
export const pageTitle = (id: string) => kept()[id]?.title ?? `Page ${id}`;

/** The linked pages, as one text for the AI (each with its title). Long pages are cut. */
export async function contextPagesText(ids: string[], max = 24_000) {
  if (!ids.length) return "";
  await refreshPages(ids);
  let left = max;
  return ids
    .map((id) => kept()[id])
    .filter((p): p is ContextPage => !!p?.text)
    .map((p) => {
      const text = p.text.length > left ? `${p.text.slice(0, Math.max(0, left))}\n…[cut: the contexts are too long]` : p.text;
      left -= text.length;
      return `### Context "${p.title}"\n${text}`;
    })
    .join("\n\n");
}
