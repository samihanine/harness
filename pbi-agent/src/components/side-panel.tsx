import { createContext, useContext, useEffect, useRef, useState } from "react";
import type { ReactNode } from "react";
import { AGENT_URL, connectAgent } from "@/agent/server";
import { cn } from "@/lib/utils";

export type SideTab = { id: string; label: string };

const Context = createContext<{ set: (tab: SideTab | null) => void; target: HTMLElement | null }>({ set: () => {}, target: null });

/**
 * Adds a tab next to the agent in the side panel while the calling page is open.
 * Returns the element to render the tab content into (with createPortal).
 */
export function useSideTab(id: string, label: string) {
  const { set, target } = useContext(Context);
  useEffect(() => {
    set({ id, label });
    return () => set(null);
  }, [set, id, label]);
  return target;
}

const MIN = 320;
const MAX = 720;

/** Left panel: the agent (iframe, always mounted) and the page's own tab if any. */
export function SidePanel({ open, children }: { open: boolean; children: ReactNode }) {
  const frame = useRef<HTMLIFrameElement>(null);
  const [tab, setTab] = useState<SideTab | null>(null);
  const [target, setTarget] = useState<HTMLElement | null>(null);
  const [active, setActive] = useState("agent");
  const [width, setWidth] = useState(() => Number(localStorage.getItem("pbi-agent:panel-width")) || 400);
  const [dragging, setDragging] = useState(false);

  useEffect(() => {
    void connectAgent(frame.current!.contentWindow!);
  }, []);
  useEffect(() => localStorage.setItem("pbi-agent:panel-width", String(width)), [width]);
  const current = tab && active === tab.id ? tab.id : "agent";

  const drag = (event: React.PointerEvent) => {
    const start = event.clientX;
    const initial = width;
    setDragging(true);
    const move = (e: PointerEvent) => setWidth(Math.min(MAX, Math.max(MIN, initial + e.clientX - start)));
    const up = () => {
      setDragging(false);
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  return (
    <Context.Provider value={{ set: setTab, target }}>
      <div className="flex min-h-0 flex-1">
        <aside className={cn("relative flex shrink-0 flex-col border-r bg-background", !open && "hidden")} style={{ width }}>
          {tab && (
            <div className="flex h-9 shrink-0 items-center gap-1 border-b px-2">
              {[{ id: "agent", label: "Agent" }, tab].map((t) => (
                <button
                  key={t.id}
                  type="button"
                  onClick={() => setActive(t.id)}
                  className={cn(
                    "rounded-md px-2 py-1 text-[12px] text-muted-foreground transition-colors hover:text-foreground",
                    current === t.id && "bg-muted text-foreground",
                  )}
                >
                  {t.label}
                </button>
              ))}
            </div>
          )}
          <iframe ref={frame} title="Agent" src={AGENT_URL} className={cn("min-h-0 w-full flex-1", current !== "agent" && "hidden", dragging && "pointer-events-none")} />
          <div ref={setTarget} className={cn("min-h-0 flex-1 overflow-y-auto", current === "agent" && "hidden")} />
          <div onPointerDown={drag} className="absolute inset-y-0 -right-1 z-10 w-2 cursor-col-resize transition-colors hover:bg-foreground/10" />
        </aside>
        <main className="min-w-0 flex-1 overflow-auto">{children}</main>
      </div>
    </Context.Provider>
  );
}
