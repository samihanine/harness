/** An embedded Power BI report filling its area (loading and errors shown over it). */
import type { Ref } from "@shared/types";
import { Loader2Icon } from "lucide-react";
import { ErrorBox } from "@/components/common";
import type { useEmbed } from "@/features/powerbi/embed";

export function ReportFrame({ embed, refOf }: { embed: ReturnType<typeof useEmbed>; refOf?: Ref }) {
  return (
    <div className="relative min-h-0 flex-1 p-3">
      <div ref={embed.element} className="h-full w-full overflow-hidden rounded-2xl bg-muted/40 [&_iframe]:border-0" />
      {refOf && embed.loading && (
        <div className="absolute inset-0 flex items-center justify-center gap-2 text-[13px] text-muted-foreground">
          <Loader2Icon className="size-4 animate-spin" /> Loading the report…
        </div>
      )}
      {embed.error && (
        <div className="absolute inset-x-6 top-6">
          <ErrorBox>{embed.error}</ErrorBox>
        </div>
      )}
    </div>
  );
}
