import { useState } from "react";
import { Mascot, type MascotState } from "./mascot";
import type { Outfit } from "./outfits";

/**
 * The mascot when nothing asks it to act: asleep, awake while the mouse is over it (or `awake`, e.g.
 * the mouse over its card). `state`: what it does when busy (thinking, working, talking), shown
 * whatever the mouse does.
 */
export function WakingMascot({ outfit, state, size, awake, className }: { outfit: Outfit; state?: MascotState; size: number; awake?: boolean; className?: string }) {
  const [hover, setHover] = useState(false);
  return (
    <span className={className} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}>
      <Mascot outfit={outfit} state={state ?? (hover || awake ? "idle" : "sleeping")} size={size} />
    </span>
  );
}
