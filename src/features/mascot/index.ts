/**
 * The mascot of the app (src/features/mascot): the drawing and its animations (mascot.tsx), its outfits
 * (outfits.tsx), and which outfit and state the assistant shows — an outfit per screen, a state from
 * what the agent is doing.
 */
import type { Status } from "@/agent/loop";
import type { Outfit } from "./outfits";
import type { MascotState } from "./mascot";

export { Mascot, STATES, type MascotState } from "./mascot";
export { WakingMascot } from "./waking";
export { OUTFITS, OUTFIT_ALIASES, type Outfit } from "./outfits";

export const OUTFIT_LABELS: Record<Outfit, string> = { diver: "Diver", trainer: "Trainer", builder: "Builder", datascientist: "Data Scientist", presenter: "Presenter" };
export const STATE_LABELS: Record<MascotState, string> = { idle: "Idle", thinking: "Thinking", talking: "Talking", working: "Working", sleeping: "Sleeping" };

/** The outfit of a screen's assistant, from its chat scope ("viewer:<id>" → trainer); the diver elsewhere. */
export function outfitOf(scope: string): Outfit {
  const screen = scope.split(":")[0];
  return ({ viewer: "trainer", editor: "builder", research: "datascientist", slides: "presenter" } as Record<string, Outfit>)[screen] ?? "diver";
}

/** While the agent runs: working when it changes things (the implementer, a save), else thinking. */
export const runningState = (status?: Status): MascotState => (status && /working|saving/i.test(status.label) ? "working" : "thinking");
