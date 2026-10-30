/**
 * Ratings of the AI's work, optional: thumbs up / down on an answer or on a whole conversation, with an
 * optional note. Kept with the conversation (downloaded with it, listed in Settings › Conversations).
 */
import { useState } from "react";
import { ThumbsDownIcon, ThumbsUpIcon } from "lucide-react";
import { Button } from "./ui/button";
import { Popover, PopoverAnchor, PopoverContent } from "./ui/popover";
import { Textarea } from "./ui/textarea";
import { find, patch, useCollection, type Rating } from "@/lib/store";
import { cn } from "@/lib/utils";

/** Writes a rating (undefined removes it): of a message when `message` is given, else of the conversation. */
async function rate(conversation: string, message: string | undefined, rating: Rating | undefined) {
  const c = await find("conversations", conversation);
  if (!c) return;
  if (!message) return patch("conversations", conversation, { rating });
  const ratings = { ...(c.ratings ?? {}) };
  if (rating) ratings[message] = rating;
  else delete ratings[message];
  return patch("conversations", conversation, { ratings });
}

/** The two thumbs; a click rates (again: removes) and offers a note. `label`: what is rated, for screen readers. */
export function RateButtons({ conversation, message, label, className }: { conversation: string; message?: string; label: string; className?: string }) {
  const c = useCollection("conversations").find((x) => x.id === conversation);
  const current = message ? c?.ratings?.[message] : c?.rating;
  const [noting, setNoting] = useState(false);
  const [note, setNote] = useState("");
  if (!c) return null;
  const pick = (score: Rating["score"]) => {
    if (current?.score === score) return void rate(conversation, message, undefined);
    void rate(conversation, message, { score, note: current?.note, at: Date.now() });
    setNote(current?.note ?? "");
    setNoting(true);
  };
  const thumb = (score: Rating["score"], Icon: typeof ThumbsUpIcon, text: string) => (
    <button
      type="button"
      aria-label={`${text}: ${label}`}
      aria-pressed={current?.score === score}
      title={current?.score === score ? `${text}${current.note ? ` — ${current.note}` : ""} (click again to remove)` : text}
      onClick={() => pick(score)}
      className={cn("icon-btn size-6", current?.score === score ? (score === "good" ? "text-success" : "text-destructive") : "text-muted-foreground/70")}
    >
      <Icon className={cn("size-3.5", current?.score === score && "fill-current")} />
    </button>
  );
  return (
    <Popover open={noting} onOpenChange={setNoting}>
      <PopoverAnchor asChild>
        <span className={cn("inline-flex items-center gap-0.5", className)}>
          {thumb("good", ThumbsUpIcon, "Good")}
          {thumb("bad", ThumbsDownIcon, "Not good")}
        </span>
      </PopoverAnchor>
      <PopoverContent align="start" className="flex w-72 flex-col gap-2 p-3">
        <span className="text-[12.5px] font-medium">{current?.score === "bad" ? "What went wrong? (optional)" : "A note? (optional)"}</span>
        <Textarea rows={3} className="text-[13px]" value={note} onChange={(e) => setNote(e.target.value)} autoFocus />
        <div className="flex justify-end gap-2">
          <Button size="sm" variant="ghost" onClick={() => setNoting(false)}>
            Skip
          </Button>
          <Button size="sm" onClick={() => (current && void rate(conversation, message, { ...current, note: note.trim() || undefined }), setNoting(false))}>
            Save
          </Button>
        </div>
      </PopoverContent>
    </Popover>
  );
}
