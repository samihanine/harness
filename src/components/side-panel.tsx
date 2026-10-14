import { useState, type PointerEvent, type ReactNode } from "react";
import { PanelLeftCloseIcon, PanelLeftOpenIcon } from "lucide-react";
import { last } from "@/lib/last";

const MIN = 280;
const MAX = 760;
const clamp = (w: number) => Math.round(Math.min(Math.max(w, MIN), Math.min(MAX, window.innerWidth - 320)));

/**
 * Left panel with tabs, which can be folded and resized (drag its right edge; width remembered in this
 * browser); every tab stays mounted (the chat keeps its state).
 */
/** `active` / `onActive`: the screen chooses the tab (e.g. a click on a row shows the Row tab); otherwise the panel does. */
export function SidePanel({ tabs, width: initial = 380, active: chosen, onActive }: { tabs: { id: string; label: string; content: ReactNode }[]; width?: number; active?: string; onActive?: (id: string) => void }) {
  const [own, setOwn] = useState(tabs[0]?.id);
  const active = chosen && tabs.some((t) => t.id === chosen) ? chosen : own;
  const setActive = (id: string) => (onActive ? onActive(id) : setOwn(id));
  const [open, setOpen] = useState(() => last.get("panel") !== "closed");
  const [width, setWidth] = useState(() => Number(last.get("panel-width")) || initial);
  const [dragging, setDragging] = useState(false);
  const toggle = (next: boolean) => {
    last.set("panel", next ? "open" : "closed");
    setOpen(next);
  };
  const resize = (e: PointerEvent<HTMLDivElement>) => {
    e.preventDefault();
    const start = e.clientX - width;
    let next = width;
    const move = (m: globalThis.PointerEvent) => setWidth((next = clamp(m.clientX - start)));
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setDragging(false);
      last.set("panel-width", String(next));
    };
    setDragging(true);
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };
  if (tabs.length === 0) return null;
  if (!open)
    return (
      <aside className="flex w-10 shrink-0 flex-col items-center gap-1 border-r bg-card py-2">
        <button type="button" className="icon-btn" title="Open the panel" onClick={() => toggle(true)}>
          <PanelLeftOpenIcon />
        </button>
        {tabs.map((t) => (
          <button key={t.id} type="button" title={t.label} onClick={() => (setActive(t.id), toggle(true))} className="w-8 truncate rounded-md py-1 text-[10px] text-muted-foreground hover:bg-muted [writing-mode:vertical-rl]">
            {t.label}
          </button>
        ))}
        {/* Tabs stay mounted while folded. */}
        <div className="hidden">{tabs.map((t) => <div key={t.id}>{t.content}</div>)}</div>
      </aside>
    );
  return (
    <aside style={{ width }} className="relative flex min-h-0 shrink-0 flex-col border-r bg-card">
      <div
        role="separator"
        aria-orientation="vertical"
        title="Drag to resize (double-click: default width)"
        onPointerDown={resize}
        onDoubleClick={() => (setWidth(initial), last.set("panel-width", String(initial)))}
        className={`absolute inset-y-0 -right-1 z-20 w-2 cursor-col-resize transition-colors hover:bg-ring/30 ${dragging ? "bg-ring/40" : ""}`}
      />
      {/* While dragging, embedded reports (iframes) must not swallow the pointer. */}
      {dragging && <div className="fixed inset-0 z-50 cursor-col-resize" />}
      <div className="flex h-12 shrink-0 items-center border-b px-2">
        {tabs.map((t) => (
          <button key={t.id} type="button" className="tab" data-active={t.id === active} onClick={() => setActive(t.id)}>
            {t.label}
          </button>
        ))}
        <button type="button" className="icon-btn ml-1" title="Fold the panel" onClick={() => toggle(false)}>
          <PanelLeftCloseIcon />
        </button>
      </div>
      {tabs.map((t) => (
        <div key={t.id} className={t.id === active ? "flex min-h-0 flex-1 flex-col overflow-auto" : "hidden"}>
          {t.content}
        </div>
      ))}
    </aside>
  );
}

export function Empty({ children }: { children: ReactNode }) {
  return <p className="m-auto max-w-72 p-6 text-center text-[12px] text-muted-foreground">{children}</p>;
}
