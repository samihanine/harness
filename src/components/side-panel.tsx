import { useState, type ReactNode } from "react";

/** Left panel with tabs; every tab stays mounted (the chat keeps its state). */
export function SidePanel({ tabs, width = 380 }: { tabs: { id: string; label: string; content: ReactNode }[]; width?: number }) {
  const [active, setActive] = useState(tabs[0].id);
  return (
    <aside style={{ width }} className="flex min-h-0 shrink-0 flex-col border-r bg-card">
      <div className="flex border-b px-2">
        {tabs.map((t) => (
          <button key={t.id} type="button" className="tab" data-active={t.id === active} onClick={() => setActive(t.id)}>
            {t.label}
          </button>
        ))}
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
