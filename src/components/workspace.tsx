/**
 * Every screen: the assistant on the left (foldable, its width dragged by its edge), and a card on the right
 * with the screen's title and chosen item, its tabs as pills, its actions, and the tab's content.
 */
import { cloneElement, isValidElement, useRef, type ReactNode } from "react";
import { PanelLeftOpenIcon } from "lucide-react";
import { Pills } from "./common";
import { useLocal } from "@/lib/store";
import { cn } from "@/lib/utils";
import type { ScreenId } from "@shared/types";

export type Tab<T extends string> = { id: T; label: string; icon?: ReactNode; count?: number };

export function Workspace<T extends string>(o: {
  screen: ScreenId | "settings";
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
  // The assistant's width (px, every screen), dragged by the handle between the two cards.
  const [width, setWidth] = useLocal<number | null>("chat-width", null);
  const panel = useRef<HTMLDivElement>(null);
  const drag = (e: React.PointerEvent) => {
    e.preventDefault();
    const start = e.clientX;
    const from = panel.current?.getBoundingClientRect().width ?? 400;
    const move = (ev: PointerEvent) => {
      const next = Math.round(Math.min(window.innerWidth * 0.6, Math.max(300, from + ev.clientX - start)));
      if (panel.current) panel.current.style.width = `${next}px`;
    };
    const up = (ev: PointerEvent) => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      document.body.style.cursor = "";
      setWidth(Math.round(Math.min(window.innerWidth * 0.6, Math.max(300, from + ev.clientX - start))));
    };
    document.body.style.cursor = "col-resize";
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };
  return (
    <div className="flex h-full gap-1">
      {o.chat &&
        (folded ? (
          <button onClick={() => setFolded(false)} className="mr-2 flex w-12 shrink-0 flex-col items-center gap-3 rounded-3xl border bg-card py-4 text-muted-foreground hover:text-foreground" aria-label="Open the assistant">
            <PanelLeftOpenIcon className="size-5" />
            <span className="text-[12px] font-medium [writing-mode:vertical-rl]">Assistant</span>
          </button>
        ) : (
          <>
            <div ref={panel} style={width ? { width } : undefined} className={cn("flex min-w-[300px] max-w-[60vw] shrink-0 flex-col overflow-hidden rounded-3xl border bg-card", !width && "w-[min(420px,38vw)]")}>{withFold(o.chat, () => setFolded(true))}</div>
            <div onPointerDown={drag} onDoubleClick={() => setWidth(null)} title="Drag to resize the assistant (double-click: default width)" className="group flex w-2 shrink-0 cursor-col-resize items-center justify-center">
              <span className="h-10 w-1 rounded-full bg-border transition-colors group-hover:bg-foreground/30" />
            </div>
          </>
        ))}
      <div className="flex min-w-0 flex-1 flex-col overflow-hidden rounded-3xl border bg-card">
        {/* As tall as the chat's header: the two lines meet. */}
        <header className="flex min-h-16 flex-wrap items-center gap-x-3 gap-y-1 border-b px-5 py-2 xl:flex-nowrap">
          <h1 className="shrink-0 text-[15px] font-semibold tracking-tight">{o.title}</h1>
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
