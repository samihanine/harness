/**
 * The brand letter of the deck (theme.brand, else the title's first letter), in a Didone (Bodoni
 * Moda): its outline draws itself, it fills in, then floats very slightly. Shown at once when the
 * slide is still (thumbnails, PDF).
 */
import { motion } from "motion/react";

export const MONOGRAM_FONT = "Bodoni Moda";

export function Monogram({ letter, still = false, delay = 0, className = "" }: { letter: string; still?: boolean; delay?: number; className?: string }) {
  const glyph = (letter.trim() || "·").slice(0, 2);
  // Outline length of one Didone capital at this size (enough for the dash to cover it).
  const length = 260 * glyph.length;
  return (
    <motion.svg
      viewBox={`0 0 ${80 * glyph.length} 80`}
      className={`overflow-visible ${className}`}
      style={{ color: "var(--s-primary)" }}
      role="img"
      aria-label={glyph}
      animate={still ? undefined : { y: [0, -1.5, 0] }}
      transition={{ duration: 6, ease: "easeInOut", repeat: Infinity, delay: delay + 2.4 }}
    >
      <motion.text
        x={40 * glyph.length}
        y={66}
        textAnchor="middle"
        fontSize={84}
        style={{ fontFamily: `"${MONOGRAM_FONT}", Didot, "Bodoni 72", serif`, fontVariationSettings: '"opsz" 96' }}
        fill="currentColor"
        stroke="currentColor"
        strokeWidth={0.6}
        strokeDasharray={length}
        initial={still ? false : { strokeDashoffset: length, fillOpacity: 0 }}
        animate={{ strokeDashoffset: 0, fillOpacity: 1 }}
        transition={{ strokeDashoffset: { duration: 2.2, ease: [0.45, 0, 0.2, 1], delay }, fillOpacity: { duration: 1.4, delay: delay + 1.3 } }}
      >
        {glyph}
      </motion.text>
    </motion.svg>
  );
}
