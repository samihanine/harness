/**
 * Every screen: the assistant on the left (foldable), and a card on the right with the screen's title
 * and chosen item, its tabs as pills, its actions, and the tab's content.
 */
import { cloneElement, isValidElement, type ReactNode } from "react";
import { PanelLeftOpenIcon } from "lucide-react";
import { Pills } from "./common";
import { useLocal } from "@/lib/store";
import { cn } from "@/lib/utils";
import type { ScreenId } from "@shared/types";

export type Tab<T extends string> = { id: T; label: string; icon?: ReactNode; count?: number };

export function Workspace<T extends string>(o: {
  screen: ScreenId;
  title: string;
  /** The chosen item (a picker) next to the title. */
  item?: ReactNode;
  actions?: ReactNode;
  /** On the right of the tabs (e.g. a mode switch). */
  aside?: ReactNode;
  tabs?: Tab<T>[];
  tab?: T;
  onTab?: (t: T) => void;
  chat?: ReactNode;
  children: ReactNode;
  /** Content scrolls (false: the tab fills the card, e.g. an embedded report). */
  scroll?: boolean;
}) {
  const [folded, setFolded] = useLocal(`chat-folded:${o.screen}`, false);
  return (
    <div className="flex h-full gap-3">
      {o.chat &&
        (folded ? (
          <button onClick={() => setFolded(false)} className="flex w-12 shrink-0 flex-col items-center gap-3 rounded-3xl border bg-card py-4 text-muted-foreground hover:text-foreground" aria-label="Open the assistant">
            <PanelLeftOpenIcon className="size-5" />
            <span className="text-[12px] font-medium [writing-mode:vertical-rl]">Assistant</span>
          </button>
        ) : (
          <div className="flex w-[min(420px,38vw)] min-w-[320px] shrink-0 flex-col overflow-hidden rounded-3xl border bg-card">{withFold(o.chat, () => setFolded(true))}</div>
        ))}
      <div className="flex min-w-0 flex-1 flex-col overflow-hidden rounded-3xl border bg-card">
        <header className="flex flex-wrap items-center gap-x-4 gap-y-2 border-b px-7 pt-5 pb-4 xl:flex-nowrap">
          <h1 className="shrink-0 text-[26px] font-light tracking-tight text-foreground/90">{o.title}</h1>
          {o.item && (
            <>
              <span className="size-1.5 shrink-0 rounded-full bg-border" />
              <div className="min-w-0 flex-1 [&>button]:max-w-full">{o.item}</div>
            </>
          )}
          {!o.item && <div className="flex-1" />}
          <div className="flex shrink-0 flex-wrap items-center gap-2">{o.actions}</div>
        </header>
        {((o.tabs && o.tabs.length > 1 && o.tab && o.onTab) || o.aside) && (
          <div className="flex flex-wrap items-center gap-2 border-b px-5 py-2.5">
            {o.tabs && o.tab && o.onTab && <Pills value={o.tab} onChange={o.onTab} options={o.tabs} />}
            <div className="flex-1" />
            {o.aside}
          </div>
        )}
        <div className={cn("min-h-0 flex-1", o.scroll === false ? "flex flex-col" : "overflow-auto")}>{o.children}</div>
      </div>
    </div>
  );
}

/** The chat panel gets the "fold" action as a prop. */
function withFold(chat: ReactNode, fold: () => void) {
  return isValidElement(chat) ? cloneElement(chat as React.ReactElement<{ onFold?: () => void }>, { onFold: fold }) : chat;
}
