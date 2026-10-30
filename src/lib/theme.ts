/** Light / dark / system theme of this browser: a `.dark` class on <html>, remembered in localStorage. */
import { useEffect, useState } from "react";

export type Theme = "light" | "dark" | "system";
const KEY = "theme";
const listeners = new Set<() => void>();

const stored = (): Theme => {
  try {
    const t = localStorage.getItem(KEY);
    return t === "dark" || t === "light" ? t : "system";
  } catch {
    return "system";
  }
};
const dark = (t: Theme) => t === "dark" || (t === "system" && typeof matchMedia !== "undefined" && matchMedia("(prefers-color-scheme: dark)").matches);

/** Applies the theme to the page (at start and on each change). */
export function applyTheme(t: Theme = stored()) {
  if (typeof document !== "undefined") document.documentElement.classList.toggle("dark", dark(t));
}

export function setTheme(t: Theme) {
  try {
    localStorage.setItem(KEY, t);
  } catch {}
  applyTheme(t);
  listeners.forEach((l) => l());
}

/** The chosen theme and the one shown (for components that draw colors themselves, e.g. charts and toasts). */
export function useTheme() {
  const [theme, set] = useState<Theme>(stored);
  useEffect(() => {
    const update = () => set(stored());
    listeners.add(update);
    const media = matchMedia("(prefers-color-scheme: dark)");
    const system = () => stored() === "system" && (applyTheme(), update());
    media.addEventListener("change", system);
    return () => {
      listeners.delete(update);
      media.removeEventListener("change", system);
    };
  }, []);
  return { theme, resolved: dark(theme) ? ("dark" as const) : ("light" as const), setTheme };
}
