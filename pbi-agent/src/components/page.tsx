import type { ReactNode } from "react";

/** Page frame: title row with actions, then content. */
export function Page({ title, subtitle, actions, children }: { title: ReactNode; subtitle?: ReactNode; actions?: ReactNode; children: ReactNode }) {
  return (
    <div className="mx-auto flex max-w-4xl flex-col gap-5 px-6 py-8 animate-in fade-in-0">
      <div className="flex items-end gap-3">
        <div className="min-w-0 flex-1">
          <h1 className="truncate text-[15px] font-medium tracking-tight">{title}</h1>
          {subtitle && <p className="text-[12px] text-muted-foreground">{subtitle}</p>}
        </div>
        {actions}
      </div>
      {children}
    </div>
  );
}

export function EmptyState({ icon, title, children }: { icon: ReactNode; title: string; children?: ReactNode }) {
  return (
    <div className="flex flex-col items-center gap-2 rounded-xl border border-dashed px-6 py-12 text-center">
      <div className="text-muted-foreground [&_svg]:size-5">{icon}</div>
      <p className="font-medium">{title}</p>
      {children && <div className="text-[12px] text-muted-foreground">{children}</div>}
    </div>
  );
}

export const ago = (time: number) => {
  const minutes = Math.round((Date.now() - time) / 60_000);
  if (minutes < 1) return "just now";
  if (minutes < 60) return `${minutes} min ago`;
  if (minutes < 1440) return `${Math.round(minutes / 60)} h ago`;
  return new Date(time).toLocaleDateString();
};
