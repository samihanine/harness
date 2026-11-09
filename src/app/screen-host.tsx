/**
 * A feature's screen in its route: loaded on first use (its own chunk; one file still in the HTML
 * export), drawn again in place when an option changes (useConfig), never mounted again.
 */
import { Suspense } from "react";
import { ScreenFallback } from "@/components/workspace";
import { useConfig } from "@/config";
import type { Feature } from "@/lib/features";

export function ScreenHost({ feature }: { feature: Feature }) {
  useConfig();
  return (
    <Suspense fallback={<ScreenFallback />}>
      <feature.Screen />
    </Suspense>
  );
}
