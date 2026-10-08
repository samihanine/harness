import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { createPortal } from "react-dom";

/** Minimal popover: fixed under its trigger (not clipped by scrolling tables), closed by outside click / Escape. */
export function Popover({ trigger, children, className = "", width = 240 }: { trigger: (open: () => void) => ReactNode; children: (close: () => void) => ReactNode; className?: string; width?: number }) {
  const anchor = useRef<HTMLSpanElement>(null);
  const panel = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState<{ top: number; left: number } | null>(null);
  const close = () => setPos(null);
  const open = () => {
    const r = anchor.current!.getBoundingClientRect();
    setPos({ top: r.bottom + 4, left: Math.min(r.left, window.innerWidth - width - 8) });
  };
  useLayoutEffect(() => {
    // Keep it on screen vertically.
    const el = panel.current;
    if (!el || !pos) return;
    const top = Math.max(8, Math.min(pos.top, window.innerHeight - el.offsetHeight - 8));
    if (top !== pos.top) setPos({ ...pos, top }); // once: the next pass computes the same top
  }, [pos]);
  useEffect(() => {
    if (!pos) return;
    const outside = (e: MouseEvent) => !panel.current?.contains(e.target as Node) && !anchor.current?.contains(e.target as Node) && close();
    const escape = (e: KeyboardEvent) => e.key === "Escape" && close();
    document.addEventListener("mousedown", outside);
    document.addEventListener("keydown", escape);
    return () => {
      document.removeEventListener("mousedown", outside);
      document.removeEventListener("keydown", escape);
    };
  }, [pos]);
  return (
    <>
      <span ref={anchor} className="block w-full">
        {trigger(open)}
      </span>
      {pos &&
        createPortal(
          <div ref={panel} style={{ top: pos.top, left: pos.left, width }} className={`fixed z-50 flex flex-col gap-2 rounded-lg border bg-card p-2 shadow-lg ${className}`}>
            {children(close)}
          </div>,
          document.body,
        )}
    </>
  );
}
