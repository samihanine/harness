import { useState, type PointerEvent, type ReactNode } from "react";
import { PanelLeftCloseIcon, PanelLeftOpenIcon, type LucideIcon } from "lucide-react";
import { last } from "@/lib/last";

const MIN = 300;
const MAX = 760;
const clamp = (w: number) => Math.round(Math.min(Math.max(w, MIN), Math.min(MAX, window.innerWidth - 420)));

export type PanelTab = { id: string; label: string; icon?: LucideIcon; content: ReactNode };

/**
 * Left panel of a screen, which can be folded and resized (drag its right edge; width remembered in
 * this browser). Several tabs (the report viewer) show as tabs; one tab shows its name. Every tab
 * stays mounted (the chat keeps its state). `active` / `onActive`: the screen may choose the tab.
 */
export function SidePanel({ tabs, width: initial = 400, active: chosen, onActive }: { tabs: PanelTab[]; width?: number; active?: string; onActive?: (id: string) => void }) {
  const [own, setOwn] = useState(tabs[0]?.id);
  const active = chosen && tabs.some((t) => t.id === chosen) ? chosen : tabs.some((t) => t.id === own) ? own : tabs[0]?.id;
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
      <aside className="flex w-12 shrink-0 flex-col items-center gap-1 border-r bg-card">
        <div className="flex h-16 w-full shrink-0 items-center justify-center border-b">
          <button type="button" className="icon-btn" title="Open the panel" onClick={() => toggle(true)}>
            <PanelLeftOpenIcon />
          </button>
        </div>
        {tabs.map((t) => (
          <button key={t.id} type="button" title={t.label} onClick={() => (setActive(t.id), toggle(true))} className="icon-btn">
            {t.icon ? <t.icon /> : <span className="text-[10px]">{t.label.slice(0, 2)}</span>}
          </button>
        ))}
        {/* Tabs stay mounted while folded. */}
        <div className="hidden">{tabs.map((t) => <div key={t.id}>{t.content}</div>)}</div>
      </aside>
    );
  const single = tabs.length === 1;
  const SingleIcon = tabs[0].icon;
  return (
    <aside style={{ width }} className="relative flex min-h-0 shrink-0 flex-col border-r bg-card">
      <div
        role="separator"
        aria-orientation="vertical"
        title="Drag to resize (double-click: default width)"
        onPointerDown={resize}
        onDoubleClick={() => (setWidth(initial), last.set("panel-width", String(initial)))}
        className={`absolute inset-y-0 -right-1 z-20 w-2 cursor-col-resize transition-colors hover:bg-primary/20 ${dragging ? "bg-primary/30" : ""}`}
      />
      {/* While dragging, embedded reports (iframes) must not swallow the pointer. */}
      {dragging && <div className="fixed inset-0 z-50 cursor-col-resize" />}
      <div className="flex h-16 shrink-0 items-center gap-5 border-b px-4">
        {single ? (
          <span className="flex items-center gap-2 text-[15px] font-semibold [&_svg]:size-4 [&_svg]:text-primary">
            {SingleIcon && <SingleIcon />} {tabs[0].label}
          </span>
        ) : (
          tabs.map((t) => (
            <button key={t.id} type="button" className="tab" data-active={t.id === active} onClick={() => setActive(t.id)}>
              {t.icon && <t.icon />} {t.label}
            </button>
          ))
        )}
        <button type="button" className="icon-btn ml-auto" title="Fold the panel" onClick={() => toggle(false)}>
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
  return <p className="m-auto max-w-80 p-8 text-center text-[13.5px] leading-relaxed text-muted-foreground">{children}</p>;
}

/**
 * Header of a screen's main area, on one line, as high as the panel's: a breadcrumb (screen / the
 * chosen item, a picker), its tools (tabs, search…) and actions on the right; `below`: a second line
 * (e.g. the Excel editor's filters).
 */
export function PageHeader({ screen, picker, actions, children, below }: { screen: string; picker?: ReactNode; actions?: ReactNode; children?: ReactNode; below?: ReactNode }) {
  return (
    <div className="shrink-0 border-b bg-card">
      <div className="flex h-16 items-center gap-2 px-5">
        <span className="shrink-0 text-[15px] text-muted-foreground">{screen}</span>
        {picker && <span className="shrink-0 text-[15px] text-muted-foreground/60">/</span>}
        {picker}
        {children}
        <span className="ml-auto" />
        {actions}
      </div>
      {below && <div className="flex min-h-14 flex-wrap items-center gap-2 border-t px-5 py-2">{below}</div>}
    </div>
  );
}
