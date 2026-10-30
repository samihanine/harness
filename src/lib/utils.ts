import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

/** Class names joined, Tailwind conflicts resolved (the last one wins). */
export const cn = (...inputs: ClassValue[]) => twMerge(clsx(inputs));
