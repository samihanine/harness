/** The deck's markdown, edited slide by slide (with its outline) or whole, with the layouts at hand. */
import { useMemo, useState } from "react";
import { LAYOUTS, joinSlides, parseDeck, plain, slideSources } from "@shared/deck";
import { Pills } from "@/components/common";
import { cn } from "@/lib/utils";

export function DeckEditor({ markdown, onChange }: { markdown: string; onChange: (md: string) => void }) {
  const [mode, setMode] = useState<"slide" | "deck">("slide");
  const [index, setIndex] = useState(0);
  const { head, slides } = useMemo(() => slideSources(markdown), [markdown]);
  const deck = useMemo(() => parseDeck(markdown), [markdown]);
  const i = Math.min(index, Math.max(0, slides.length - 1));
  return (
    <div className="flex h-full flex-col gap-3 p-5">
      <div className="flex items-center gap-3">
        <Pills value={mode} onChange={setMode} options={[{ id: "slide", label: "Slide by slide" }, { id: "deck", label: "Whole deck" }]} />
        <div className="flex-1" />
        <details className="relative">
          <summary className="chip list-none">Layouts</summary>
          <div className="absolute right-0 z-10 mt-2 w-[420px] rounded-2xl border bg-popover p-4 text-[12.5px] shadow-lg">
            {Object.entries(LAYOUTS).map(([k, v]) => (
              <div key={k} className="mb-1.5"><span className="font-mono font-semibold">{k}</span> — <span className="text-muted-foreground">{v}</span></div>
            ))}
          </div>
        </details>
      </div>
      {mode === "deck" ? (
        <textarea value={markdown} onChange={(e) => onChange(e.target.value)} spellCheck={false} className="min-h-0 flex-1 rounded-2xl border bg-muted/30 p-4 font-mono text-[12.5px] leading-relaxed outline-none focus:border-foreground/30" />
      ) : (
        <div className="flex min-h-0 flex-1 gap-3">
          <div className="flex w-56 shrink-0 flex-col gap-1 overflow-y-auto">
            {deck.slides.map((s, n) => (
              <button key={n} onClick={() => setIndex(n)} className={cn("rounded-xl px-3 py-2 text-left text-[12.5px]", n === i ? "bg-muted font-medium" : "hover:bg-muted/60")}>
                <div className="text-[11px] text-muted-foreground">{n + 1} · {s.layout}</div>
                <div className="truncate">{s.title ? plain(s.title) : "(no title)"}</div>
              </button>
            ))}
            <button onClick={() => (onChange(joinSlides(head, [...slides, "# New slide\n\n- A point"])), setIndex(slides.length))} className="rounded-xl border border-dashed px-3 py-2 text-[12.5px] text-muted-foreground hover:bg-muted">+ Slide</button>
          </div>
          <textarea
            value={slides[i] ?? ""}
            onChange={(e) => onChange(joinSlides(head, slides.map((s, n) => (n === i ? e.target.value : s))))}
            spellCheck={false}
            className="min-h-0 flex-1 rounded-2xl border bg-muted/30 p-4 font-mono text-[12.5px] leading-relaxed outline-none focus:border-foreground/30"
          />
        </div>
      )}
    </div>
  );
}
