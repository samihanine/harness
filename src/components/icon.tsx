/** A lucide icon from its name ("store", "users-round", "shield-check"…), a sparkle when unknown (diagrams, markdown icons). */
import { icons, SparklesIcon } from "lucide-react";

export const iconOf = (name: string) => icons[name.replace(/(^|[-_ ])(\w)/g, (_, __, c: string) => c.toUpperCase()) as keyof typeof icons] ?? SparklesIcon;
