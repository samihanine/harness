/** The server's settings as the browser sees them (no secret), shared by every screen. */
import { useEffect, useState } from "react";
import type { PublicSettings, SettingsPatch } from "@shared/types";
import { get, post, put } from "./api";

export type AppSettings = PublicSettings & { models: string[] };
let current: AppSettings | undefined;
let loading: Promise<AppSettings> | undefined;
const listeners = new Set<(s: AppSettings) => void>();

export function loadSettings() {
  loading ??= get<AppSettings>("/settings").then((s) => ((current = s), listeners.forEach((l) => l(s)), s));
  return loading;
}

export async function saveSettings(patch: SettingsPatch) {
  const s = await put<AppSettings>("/settings", patch);
  current = s;
  listeners.forEach((l) => l(s));
  return s;
}

export function useSettings() {
  const [s, set] = useState(current);
  useEffect(() => {
    listeners.add(set);
    void loadSettings().then(set, () => undefined);
    return () => void listeners.delete(set);
  }, []);
  return s;
}

/** Forgets the changes of a part of the settings: back to the .env values. */
export async function resetSettings(part: "general" | "ai" | "powerbi" | "confluence") {
  const s = await post<AppSettings>("/settings/reset", { part });
  current = s;
  listeners.forEach((l) => l(s));
  return s;
}
