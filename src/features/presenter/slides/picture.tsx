/**
 * The picture of a card or a row in a slide: an image (URL or "img:<id>" of the library), a lucide icon
 * by name, or the mascot in an outfit. In PowerPoint icons and the mascot go as pictures drawn before
 * the export.
 */
import type { CSSProperties } from "react";
import { icons, SparklesIcon } from "lucide-react";
import { z } from "zod";
import { Mascot, OUTFIT_ALIASES, OUTFITS, type Outfit } from "@/features/mascot";

import { pptxImage } from "./images";

/** An outfit of the mascot (older names accepted). */
export const outfitSchema = z.preprocess((v) => (typeof v === "string" ? (OUTFIT_ALIASES[v] ?? v) : v), z.enum(Object.keys(OUTFITS) as [Outfit, ...Outfit[]]));

/** A lucide icon from its name ("store", "users-round", "shield-check"…), a sparkle when unknown. */
export const iconOf = (name: string) => icons[name.replace(/(^|[-_ ])(\w)/g, (_, __, c: string) => c.toUpperCase()) as keyof typeof icons] ?? SparklesIcon;

/** A lucide icon by name, sized by its style (cqw). */
export function Icon({ name, className = "", style, strokeWidth = 1.8 }: { name?: string; className?: string; style?: CSSProperties; strokeWidth?: number }) {
  const I = iconOf(name ?? "");
  return <I className={className} style={style} strokeWidth={strokeWidth} />;
}

export const pictureArgs = {
  image: z.string().optional().describe('Picture: "img:<id>" of the library, or a public URL (a logo, a photo)'),
  icon: z.string().optional().describe("Or a lucide icon name (e.g. shield-check, database, rocket)"),
  mascot: outfitSchema.optional().describe("Or the app's mascot in an outfit: diver, trainer, builder, datascientist, presenter"),
};
export type PictureArgs = { image?: string; icon?: string; mascot?: Outfit };

/** The picture filling its box (`className` sets the box), nothing when none is given. */
export function Picture({ image, icon, mascot, className = "", cover }: PictureArgs & { className?: string; cover?: boolean }) {
  if (mascot) return <Mascot outfit={mascot} state="idle" size={200} className={`h-auto max-h-full w-full ${className}`} />;
  if (image) return <img src={image} alt="" className={`${cover ? "object-cover" : "object-contain"} ${className}`} />;
  if (icon) {
    const Icon = iconOf(icon);
    return <Icon className={className} strokeWidth={1.6} style={{ color: "var(--s-primary)" }} />;
  }
  return null;
}

/**
 * The picture in PowerPoint (`fit`: whole, or filling the box): an image, or the mascot / icon drawn
 * as a picture before the export (pptx.ts).
 */
export const pptxPicture = (p: PictureArgs, box: { x: number; y: number; w: number; h: number }, fit: "cover" | "contain" = "contain") => {
  const key = p.mascot ? `mascot:${p.mascot}` : p.image ? p.image : p.icon ? `icon:${p.icon}` : undefined;
  return key ? pptxImage(key, box, p.image ? fit : "contain") : undefined;
};
