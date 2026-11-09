/**
 * A choice as a badge: tinted background, colored dot (Excel options, statuses, filters). The colors
 * are a short fixed palette, also written into the Excel files (options' fills) by their hex value.
 */
export const COLORS = ["gray", "red", "orange", "amber", "green", "teal", "blue", "violet", "pink"] as const;
export type Color = (typeof COLORS)[number];
export const isColor = (v: unknown): v is Color => (COLORS as readonly unknown[]).includes(v);

/** Tinted background and text per color. */
export const TINTS: Record<Color, string> = {
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

/** Solid swatch per color (color pickers, dots). */
export const SWATCHES: Record<Color, string> = {
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

/** The same colors as fills of Excel cells (light, so the text stays readable) and as their font color. */
export const EXCEL_FILLS: Record<Color, { fill: string; font: string }> = {
  gray: { fill: "#E4E4E7", font: "#3F3F46" },
  red: { fill: "#FEE2E2", font: "#B91C1C" },
  orange: { fill: "#FFEDD5", font: "#C2410C" },
  amber: { fill: "#FEF3C7", font: "#92400E" },
  green: { fill: "#D1FAE5", font: "#047857" },
  teal: { fill: "#CCFBF1", font: "#0F766E" },
  blue: { fill: "#DBEAFE", font: "#1D4ED8" },
  violet: { fill: "#EDE9FE", font: "#6D28D9" },
  pink: { fill: "#FCE7F3", font: "#BE185D" },
};

export const Badge = ({ color, small, children }: { color?: Color; small?: boolean; children: string }) => (
  <span title={children} className={`badge max-w-full min-w-0 ${small ? "h-5.5 gap-1 px-2 text-[11.5px]" : ""} ${color ? TINTS[color] : "bg-red-500/10 text-destructive"}`}>
    <span className={`size-1.5 shrink-0 rounded-full ${color ? SWATCHES[color] : "bg-destructive"}`} />
    <span className="truncate">{children}</span>
  </span>
);
