import { useState, type ReactNode } from "react";
import { PanelLeftCloseIcon, PanelLeftOpenIcon } from "lucide-react";
import { last } from "@/lib/last";

/** Left panel with tabs, which can be folded; every tab stays mounted (the chat keeps its state). */
export function SidePanel({ tabs, width = 380 }: { tabs: { id: string; label: string; content: ReactNode }[]; width?: number }) {
  const [active, setActive] = useState(tabs[0]?.id);
  const [open, setOpen] = useState(() => last.get("panel") !== "closed");
  const toggle = (next: boolean) => {
    last.set("panel", next ? "open" : "closed");
    setOpen(next);
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
    <aside style={{ width }} className="flex min-h-0 shrink-0 flex-col border-r bg-card">
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
