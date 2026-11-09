/**
 * The registry of features. Each feature (src/features/<name>) describes itself in its `module.ts`
 * (its screen, address, name, picture, settings tabs, the items its conversations are about); the app
 * layer (src/app/registry.ts) registers them at start. The shared code (layout, home page, settings,
 * conversations) only reads this list: it never imports a feature, so a feature can be added or
 * removed without touching it.
 */
import type { ComponentType, LazyExoticComponent } from "react";
import type { LucideIcon } from "lucide-react";
import { SCREENS, type Screen } from "@/config";
import type { Outfit } from "@/components/mascot";
import type { Name } from "./store";

type Lazy = LazyExoticComponent<ComponentType>;

/** A tab of Settings given by a feature (e.g. Presenter › its PowerPoint template). */
export type SettingsTab = { id: string; label: string; icon: LucideIcon; Panel: Lazy };

export type Feature = {
  /** Its flag in SCREENS (config.ts), and the prefix of its conversations' scope. */
  id: Screen;
  path: string;
  label: string;
  /** What it is for, in a line (home page). */
  description: string;
  /** 0: the AI assistants, 1: the library. */
  group: 0 | 1;
  /** Its picture: the mascot in an outfit (assistants) or a logo (library: components/screen-logo.tsx). */
  picture: { outfit: Outfit } | { logo: "powerbi" | "model" | "excel" | "wiki" };
  /** The screen, loaded when first opened (each feature is its own chunk). */
  Screen: Lazy;
  /** Its URL parameter naming the chosen item (e.g. "report"): `?report=<id>&tab=<tab>`. */
  param: string;
  /** The collection its conversations are about (Conversations dialog: the item's name). */
  collection?: Name;
  settings?: SettingsTab[];
  /** Its items can be given Confluence pages as context (Confluence › Linked to); they have a name or title and contextIds. */
  linkable?: { collection: Name; title: string };
};

let registered: Feature[] = [];
export const registerFeatures = (list: Feature[]) => void (registered = list);

/** Every feature, in the order of the sidebar. */
export const features = () => registered;
/** The features shown (their flag on). */
export const shownFeatures = () => registered.filter((f) => SCREENS[f.id]);
export const featureOf = (id: string) => registered.find((f) => f.id === id);

/** Where a hidden screen (or "/" without a home page) leads: the first screen shown, else Settings. */
export const firstScreen = () => shownFeatures()[0]?.path ?? "/settings";

/** The search of a screen's route: its item and tab (strings: a numeric id such as a Confluence page's stays a string). */
const text = (v: unknown) => (v === undefined || v === null || v === "" ? undefined : String(v));
export const itemSearch =
  <K extends string>(param: K) =>
  (s: Record<string, unknown>) =>
    ({ [param]: text(s[param]), tab: text(s.tab) }) as { [P in K]?: string } & { tab?: string };
