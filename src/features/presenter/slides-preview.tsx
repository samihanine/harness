/** Pictures of the slides as PowerPoint draws them (rendered by the server through Microsoft): a grid, one slide large on click. */
import { useEffect, useRef, useState } from "react";
import { Loader2Icon, RefreshCwIcon } from "lucide-react";
import { ErrorBox } from "@/components/common";
import { Button } from "@/components/ui/button";
import { errorText, fileUrl, post } from "@/lib/api";
import { cn } from "@/lib/utils";

const cache = new Map<string, string[]>();

export function SlidesPreview({ source, cacheKey, auto = true }: { source: { markdown: string } | { fileId: string }; cacheKey: string; auto?: boolean }) {
  const [pics, setPics] = useState<string[] | undefined>(cache.get(cacheKey));
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string>();
  const [big, setBig] = useState<number>();
  const asked = useRef("");
  const render = async () => {
    setBusy(true);
    setError(undefined);
    const key = cacheKey;
    asked.current = key;
    try {
      const ids = await post<string[]>("/deck/preview", source);
      cache.set(key, ids);
      if (asked.current === key) setPics(ids);
    } catch (e) {
      if (asked.current === key) setError(errorText(e));
    } finally {
      if (asked.current === key) setBusy(false);
    }
  };
  useEffect(() => {
    const hit = cache.get(cacheKey);
    if (hit) return setPics(hit);
    if (auto) void render();
  }, [cacheKey]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <div className="flex flex-col gap-4 p-6">
      <div className="flex items-center gap-2">
        <div className="flex-1 text-[13px] text-muted-foreground">{pics ? `${pics.length} slides, as PowerPoint draws them` : "Preview drawn by PowerPoint (a few seconds)"}</div>
        <Button variant="outline" size="sm" onClick={render} disabled={busy}>
          {busy ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} {busy ? "Drawing…" : "Refresh"}
        </Button>
      </div>
      {error && <ErrorBox>{error}</ErrorBox>}
      {big !== undefined && pics?.[big] ? (
        <div className="flex flex-col gap-3">
          <img src={fileUrl(pics[big])} alt={`Slide ${big + 1}`} className="w-full rounded-2xl border shadow-sm" onClick={() => setBig(undefined)} />
          <div className="flex justify-center gap-2">
            <Button variant="outline" size="sm" disabled={big === 0} onClick={() => setBig(big - 1)}>Previous</Button>
            <Button variant="outline" size="sm" onClick={() => setBig(undefined)}>All slides</Button>
            <Button variant="outline" size="sm" disabled={big === pics.length - 1} onClick={() => setBig(big + 1)}>Next</Button>
          </div>
        </div>
      ) : (
        <div className={cn("grid gap-4 sm:grid-cols-2 2xl:grid-cols-3", busy && "opacity-60")}>
          {pics?.map((id, i) => (
            <button key={id} onClick={() => setBig(i)} className="group text-left">
              <img src={fileUrl(id)} alt={`Slide ${i + 1}`} className="w-full rounded-xl border shadow-[0_1px_4px_rgba(0,0,0,0.06)] transition-transform group-hover:-translate-y-0.5" />
              <div className="mt-1 text-[12px] text-muted-foreground">{i + 1}</div>
            </button>
          ))}
          {!pics && busy && Array.from({ length: 4 }, (_, i) => <div key={i} className="aspect-video animate-pulse rounded-xl bg-muted" />)}
        </div>
      )}
    </div>
  );
}
