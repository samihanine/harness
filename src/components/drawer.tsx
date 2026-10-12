import { useEffect, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { XIcon } from "lucide-react";

/** Side sheet over the screen (editing kept out of the reading views); closed by Escape or outside click. */
export function Drawer({ title, onClose, children, footer }: { title: string; onClose: () => void; children: ReactNode; footer?: ReactNode }) {
  useEffect(() => {
    const escape = (e: KeyboardEvent) => e.key === "Escape" && onClose();
    document.addEventListener("keydown", escape);
    return () => document.removeEventListener("keydown", escape);
  }, [onClose]);
  return createPortal(
    <div className="fixed inset-0 z-40 flex justify-end bg-black/20" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="flex h-full w-full max-w-md flex-col border-l bg-card shadow-2xl">
        <div className="flex h-12 shrink-0 items-center border-b px-4">
          <h2 className="font-medium">{title}</h2>
          <button type="button" className="icon-btn ml-auto" title="Close" onClick={onClose}>
            <XIcon />
          </button>
        </div>
        <div className="flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto p-4">{children}</div>
        {footer && <div className="flex shrink-0 items-center gap-2 border-t px-4 py-2.5">{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}

