/**
 * The Edit tab. A deck from the template: its slides on the left (pictures as PowerPoint draws them), the
 * chosen one large with its markdown to change (or the whole deck's markdown). An imported PowerPoint: its
 * slides, the chosen one large with what it holds (shapes, texts, pictures, tables).
 */
import { useEffect, useMemo, useState } from "react";
import { ArrowDownIcon, ArrowUpIcon, CopyPlusIcon, ImageIcon, Loader2Icon, RefreshCwIcon, Trash2Icon, TypeIcon } from "lucide-react";
import { joinSlides, parseDeck, plain, slideSources } from "@shared/deck";
import { Pills } from "@/components/common";
import { Button } from "@/components/ui/button";
import { fileUrl, get } from "@/lib/api";
import { cn } from "@/lib/utils";
import type { useSlides } from "./use-slides";

function Thumbs({ pics, count, index, onPick, label }: { pics?: string[]; count: number; index: number; onPick: (i: number) => void; label: (i: number) => string }) {
  return (
    <div className="flex w-52 shrink-0 flex-col gap-2 overflow-y-auto pr-1">
      {Array.from({ length: count }, (_, i) => (
        <button key={i} onClick={() => onPick(i)} className={cn("rounded-xl border-2 p-1 text-left transition-colors", i === index ? "border-primary" : "border-transparent hover:border-border")}>
          {pics?.[i] ? <img src={fileUrl(pics[i])} alt="" className="w-full rounded-lg border" /> : <div className="aspect-video rounded-lg bg-muted" />}
          <div className="mt-1 truncate px-1 text-[11.5px] text-muted-foreground">{i + 1} · {label(i)}</div>
        </button>
      ))}
    </div>
  );
}

export function EditDeck({ markdown, onChange, slides }: { markdown: string; onChange: (md: string) => void; slides: ReturnType<typeof useSlides> }) {
  const [mode, setMode] = useState<"slide" | "deck">("slide");
  const [index, setIndex] = useState(0);
  const { head, slides: sources } = useMemo(() => slideSources(markdown), [markdown]);
  const deck = useMemo(() => parseDeck(markdown), [markdown]);
  const i = Math.min(index, Math.max(0, sources.length - 1));
  const set = (list: string[], next = i) => (onChange(joinSlides(head, list)), setIndex(Math.max(0, Math.min(list.length - 1, next))));
  const stale = slides.pics && slides.pics.length !== sources.length;
  return (
    <div className="flex h-full flex-col gap-3 p-5">
      <div className="flex flex-wrap items-center gap-2">
        <Pills value={mode} onChange={setMode} options={[{ id: "slide", label: "Slide by slide" }, { id: "deck", label: "Whole deck (markdown)" }]} />
        <div className="flex-1" />
        <Button variant="outline" size="sm" onClick={slides.render} disabled={slides.busy}>{slides.busy ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Refresh the preview</Button>
      </div>
      {mode === "deck" ? (
        <textarea value={markdown} onChange={(e) => onChange(e.target.value)} spellCheck={false} className="min-h-0 flex-1 rounded-2xl border bg-muted/30 p-4 font-mono text-[12.5px] leading-relaxed outline-none focus:border-foreground/30" />
      ) : (
        <div className="flex min-h-0 flex-1 gap-4">
          <Thumbs pics={slides.pics} count={sources.length} index={i} onPick={setIndex} label={(n) => (deck.slides[n]?.title ? plain(deck.slides[n].title) : deck.slides[n]?.layout ?? "")} />
          <div className="flex min-w-0 flex-1 flex-col gap-3">
            <div className="relative">
              {slides.pics?.[i] ? <img src={fileUrl(slides.pics[i])} alt={`Slide ${i + 1}`} className={cn("w-full rounded-2xl border shadow-sm", (slides.busy || stale) && "opacity-60")} /> : <div className="aspect-video animate-pulse rounded-2xl bg-muted" />}
              {(slides.busy || stale) && <div className="absolute top-3 right-3 rounded-full bg-card px-3 py-1 text-[12px] shadow">{slides.busy ? "Drawing…" : "Changed: refresh the preview"}</div>}
            </div>
            <div className="flex items-center gap-1">
              <span className="flex-1 text-[12.5px] text-muted-foreground">Slide {i + 1} · layout {deck.slides[i]?.layout}</span>
              <Button variant="ghost" size="icon-sm" title="Move up" disabled={i === 0} onClick={() => { const l = [...sources]; [l[i - 1], l[i]] = [l[i], l[i - 1]]; set(l, i - 1); }}><ArrowUpIcon /></Button>
              <Button variant="ghost" size="icon-sm" title="Move down" disabled={i >= sources.length - 1} onClick={() => { const l = [...sources]; [l[i + 1], l[i]] = [l[i], l[i + 1]]; set(l, i + 1); }}><ArrowDownIcon /></Button>
              <Button variant="ghost" size="icon-sm" title="Duplicate" onClick={() => set([...sources.slice(0, i + 1), sources[i], ...sources.slice(i + 1)], i + 1)}><CopyPlusIcon /></Button>
              <Button variant="ghost" size="icon-sm" title="Delete" disabled={sources.length < 2} onClick={() => set(sources.filter((_, n) => n !== i))}><Trash2Icon /></Button>
            </div>
            <textarea value={sources[i] ?? ""} onChange={(e) => onChange(joinSlides(head, sources.map((s, n) => (n === i ? e.target.value : s))))} spellCheck={false} className="min-h-48 flex-1 rounded-2xl border bg-muted/30 p-4 font-mono text-[12.5px] leading-relaxed outline-none focus:border-foreground/30" />
          </div>
        </div>
      )}
    </div>
  );
}

type Shape = { id: string; name: string; kind: string; placeholder?: string; x: number; y: number; w: number; h: number; text?: string; rows?: string[][]; image?: { width: number; height: number } };
type SlideInfo = { number: number; layout: string; shapes: Shape[]; notes?: string };

export function EditPptx({ fileId, slides }: { fileId: string; slides: ReturnType<typeof useSlides> }) {
  const [structure, setStructure] = useState<SlideInfo[]>();
  const [index, setIndex] = useState(0);
  useEffect(() => void get<SlideInfo[]>(`/pptx/structure?fileId=${fileId}`).then(setStructure, () => setStructure([])), [fileId]);
  const count = structure?.length ?? slides.pics?.length ?? 0;
  const s = structure?.[index];
  return (
    <div className="flex h-full gap-4 p-5">
      <Thumbs pics={slides.pics} count={count} index={index} onPick={setIndex} label={(n) => structure?.[n]?.layout ?? ""} />
      <div className="flex min-w-0 flex-1 flex-col gap-3 overflow-y-auto">
        {slides.pics?.[index] ? <img src={fileUrl(slides.pics[index])} alt="" className="w-full rounded-2xl border shadow-sm" /> : <div className="aspect-video animate-pulse rounded-2xl bg-muted" />}
        <div className="text-[12.5px] text-muted-foreground">Slide {index + 1}{s?.layout ? ` · layout "${s.layout}"` : ""} · ask the assistant to change a text, a picture, a table, or to add a slide from the template.</div>
        <div className="flex flex-col divide-y rounded-2xl border">
          {s?.shapes.map((x) => (
            <div key={x.id} className="flex items-start gap-3 px-4 py-2 text-[12.5px]">
              {x.kind === "picture" ? <ImageIcon className="mt-0.5 size-4 text-muted-foreground" /> : <TypeIcon className="mt-0.5 size-4 text-muted-foreground" />}
              <span className="w-40 shrink-0 truncate font-medium">{x.name}</span>
              <span className="min-w-0 flex-1 text-muted-foreground">{x.text ? x.text.slice(0, 200) : x.rows ? `table ${x.rows.length}×${x.rows[0]?.length ?? 0}` : x.image ? `picture ${x.image.width}×${x.image.height}px` : x.kind}</span>
            </div>
          ))}
          {s?.notes && <div className="px-4 py-2 text-[12.5px] text-muted-foreground">Notes: {s.notes}</div>}
        </div>
      </div>
    </div>
  );
}
