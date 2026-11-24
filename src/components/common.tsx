/** Small shared pieces: empty states, sections, notices, loading. */
import type { ReactNode } from "react";
import { AlertTriangleIcon, Loader2Icon } from "lucide-react";
import { cn } from "@/lib/utils";

export function Empty({ icon, title, children, className }: { icon?: ReactNode; title: string; children?: ReactNode; className?: string }) {
  return (
    <div className={cn("flex h-full min-h-60 flex-col items-center justify-center gap-3 p-8 text-center", className)}>
      {icon && <div className="flex size-12 items-center justify-center rounded-2xl bg-muted text-muted-foreground [&_svg]:size-6">{icon}</div>}
      <div className="text-[15px] font-semibold">{title}</div>
      {children && <div className="max-w-md text-[13.5px] text-muted-foreground">{children}</div>}
    </div>
  );
}

export function Section({ title, actions, children, className }: { title?: ReactNode; actions?: ReactNode; children: ReactNode; className?: string }) {
  return (
    <section className={cn("rounded-2xl border bg-card", className)}>
      {(title || actions) && (
        <header className="flex items-center gap-2 border-b px-5 py-3">
          <div className="flex-1 text-[13.5px] font-semibold">{title}</div>
          {actions}
        </header>
      )}
      <div className="p-5">{children}</div>
    </section>
  );
}

export const Loading = ({ label = "Loading…" }: { label?: string }) => (
  <div className="flex h-full min-h-40 items-center justify-center gap-2 text-[13px] text-muted-foreground">
    <Loader2Icon className="size-4 animate-spin" /> {label}
  </div>
);

export const ErrorBox = ({ children, className }: { children: ReactNode; className?: string }) => (
  <div className={cn("flex items-start gap-2 rounded-xl border border-destructive/30 bg-destructive/5 px-4 py-3 text-[13px] text-destructive", className)}>
    <AlertTriangleIcon className="mt-0.5 size-4 shrink-0" />
    <div className="min-w-0 break-words">{children}</div>
  </div>
);

/** Pills of choices (tabs, modes). */
export function Pills<T extends string>({ value, onChange, options, className }: { value: T; onChange: (v: T) => void; options: { id: T; label: ReactNode; icon?: ReactNode; count?: number }[]; className?: string }) {
  return (
    <div className={cn("flex flex-wrap items-center gap-1.5", className)}>
      {options.map((o) => (
        <button
          key={o.id}
          onClick={() => onChange(o.id)}
          className={cn(
            "inline-flex h-10 items-center gap-2 rounded-full px-4 text-[13.5px] font-medium transition-colors [&_svg]:size-4",
            value === o.id ? "bg-muted text-foreground ring-1 ring-border" : "text-muted-foreground hover:bg-muted/60 hover:text-foreground",
          )}
        >
          {o.icon}
          {o.label}
          {o.count !== undefined && <span className="text-[12px] tabular-nums text-muted-foreground">{o.count}</span>}
        </button>
      ))}
    </div>
  );
}
