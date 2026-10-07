import type { OptionValue } from "./schema";

/** Tinted badge per option color (light and dark). */
export const TINTS: Record<OptionValue["color"], string> = {
  gray: "bg-zinc-500/10 text-zinc-700 dark:text-zinc-300",
  red: "bg-red-500/10 text-red-700 dark:text-red-300",
  orange: "bg-orange-500/10 text-orange-700 dark:text-orange-300",
  amber: "bg-amber-500/15 text-amber-800 dark:text-amber-300",
  green: "bg-emerald-500/10 text-emerald-700 dark:text-emerald-300",
  teal: "bg-teal-500/10 text-teal-700 dark:text-teal-300",
  blue: "bg-blue-500/10 text-blue-700 dark:text-blue-300",
  violet: "bg-violet-500/10 text-violet-700 dark:text-violet-300",
  pink: "bg-pink-500/10 text-pink-700 dark:text-pink-300",
};

export const DOTS: Record<OptionValue["color"], string> = {
  gray: "bg-zinc-400",
  red: "bg-red-500",
  orange: "bg-orange-500",
  amber: "bg-amber-500",
  green: "bg-emerald-500",
  teal: "bg-teal-500",
  blue: "bg-blue-500",
  violet: "bg-violet-500",
  pink: "bg-pink-500",
};

/** Option colours in the Excel file: light fill, dark text (same hues as the badges). */
export const CELL_COLORS: Record<OptionValue["color"], { fill: string; font: string }> = {
  gray: { fill: "#f4f4f5", font: "#3f3f46" },
  red: { fill: "#fee2e2", font: "#b91c1c" },
  orange: { fill: "#ffedd5", font: "#c2410c" },
  amber: { fill: "#fef3c7", font: "#92400e" },
  green: { fill: "#d1fae5", font: "#047857" },
  teal: { fill: "#ccfbf1", font: "#0f766e" },
  blue: { fill: "#dbeafe", font: "#1d4ed8" },
  violet: { fill: "#ede9fe", font: "#6d28d9" },
  pink: { fill: "#fce7f3", font: "#be185d" },
};
