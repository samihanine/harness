/**
 * A feature's screen in its route: loaded on first use (its own chunk; one file still in the HTML
 * export), drawn again in place when an option changes (useConfig), never mounted again.
 */
import { Suspense } from "react";
import { ErrorBoundary } from "@/components/error-boundary";
import { ScreenFallback } from "@/components/workspace";
import { useConfig } from "@/config";
import type { Feature } from "@/lib/features";

export function ScreenHost({ feature }: { feature: Feature }) {
  useConfig();
  return (
    // A screen that fails shows why and offers to try again; the sidebar and the other screens keep working.
    <ErrorBoundary
      resetKey={feature.id}
      fallback={(error, reset) => (
        <div className="flex h-full flex-col items-center justify-center gap-3 p-8 text-center">
          <p className="text-[15px] font-medium">{feature.label} could not be drawn</p>
          <p className="max-w-xl text-[13px] break-words text-muted-foreground">{error.message}</p>
          <button type="button" className="btn" onClick={reset}>
            Try again
          </button>
        </div>
      )}
    >
      <Suspense fallback={<ScreenFallback />}>
        <feature.Screen />
      </Suspense>
    </ErrorBoundary>
  );
}
