/**
 * The mascot of the app (src/features/mascot): the drawing and its animations (mascot.tsx), its outfits
 * (outfits.tsx), and which outfit and state the assistant shows — an outfit per screen, a state from
 * what the agent is doing.
 */
import type { Outfit } from "./outfits";
import type { MascotState } from "./mascot";

export { Mascot, STATES, type MascotState } from "./mascot";
export { WakingMascot } from "./waking";
export { OUTFITS, OUTFIT_ALIASES, type Outfit } from "./outfits";

export const OUTFIT_LABELS: Record<Outfit, string> = { diver: "Diver", trainer: "Trainer", builder: "Builder", datascientist: "Data Scientist", presenter: "Presenter" };
export const STATE_LABELS: Record<MascotState, string> = { idle: "Idle", thinking: "Thinking", talking: "Talking", working: "Working", sleeping: "Sleeping" };

/** The outfit of a screen's assistant; the diver elsewhere. */
export const outfitOf = (screen: string): Outfit => (({ trainer: "trainer", builder: "builder", scientist: "datascientist", presenter: "presenter" }) as Record<string, Outfit>)[screen] ?? "diver";
