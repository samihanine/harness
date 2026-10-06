import { useEffect, useRef, useState } from "react";
import { AGENT_URL, connectAgent } from "@/agent/server";
import { cn } from "@/lib/utils";

const MIN = 320;
const MAX = 720;

/** The agent iframe, on the right, resizable. Stays mounted when hidden (keeps its state). */
export function AgentFrame({ open }: { open: boolean }) {
  const frame = useRef<HTMLIFrameElement>(null);
  const [width, setWidth] = useState(() => Number(localStorage.getItem("excel-agent:agent-width")) || 400);
  const [dragging, setDragging] = useState(false);

  useEffect(() => {
    void connectAgent(frame.current!.contentWindow!);
  }, []);
  useEffect(() => localStorage.setItem("excel-agent:agent-width", String(width)), [width]);

  const drag = (event: React.PointerEvent) => {
    const start = event.clientX;
    const initial = width;
    setDragging(true);
    const move = (e: PointerEvent) => setWidth(Math.min(MAX, Math.max(MIN, initial + start - e.clientX)));
    const up = () => {
      setDragging(false);
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  return (
    <aside className={cn("relative shrink-0 border-l bg-background", !open && "hidden")} style={{ width }}>
      <div
        onPointerDown={drag}
        className="absolute inset-y-0 -left-1 z-10 w-2 cursor-col-resize transition-colors hover:bg-foreground/10"
      />
      <iframe ref={frame} title="Agent" src={AGENT_URL} className={cn("size-full", dragging && "pointer-events-none")} />
    </aside>
  );
}
