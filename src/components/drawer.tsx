import { useEffect, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { XIcon } from "lucide-react";

/** Side sheet over the screen (forms and details kept out of the main views); closed by Escape or outside click. */
export function Drawer({ title, onClose, children, footer, wide = false }: { title: string; onClose: () => void; children: ReactNode; footer?: ReactNode; wide?: boolean }) {
  useEffect(() => {
    const escape = (e: KeyboardEvent) => e.key === "Escape" && onClose();
    document.addEventListener("keydown", escape);
    return () => document.removeEventListener("keydown", escape);
  }, [onClose]);
  return createPortal(
    <div className="fixed inset-0 z-40 flex justify-end bg-black/25" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className={`flex h-full w-full flex-col border-l bg-card shadow-2xl ${wide ? "max-w-3xl" : "max-w-lg"}`}>
        <div className="flex h-[72px] shrink-0 items-center border-b px-6">
          <h2 className="text-[17px] font-semibold">{title}</h2>
          <button type="button" className="icon-btn ml-auto" title="Close" onClick={onClose}>
            <XIcon />
          </button>
        </div>
        <div className="flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto p-6">{children}</div>
        {footer && <div className="flex shrink-0 items-center gap-2 border-t px-6 py-4">{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}
