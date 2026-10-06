import type { ReactNode } from "react";
import { Link } from "@tanstack/react-router";
import { DatabaseIcon, LogInIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { db, useDb } from "@/db/db";
import { useAuth } from "@/lib/use-auth";
import { EmptyState } from "./page";

/** Renders its children once the database folder is open. */
export function NeedsDb({ children }: { children: ReactNode }) {
  const { state } = useDb();
  if (state.status === "ready") return <>{children}</>;
  if (state.status === "loading") return <p className="shimmer p-8 text-[12px]">Opening the database…</p>;
  return (
    <div className="mx-auto max-w-lg p-8">
      <EmptyState icon={<DatabaseIcon />} title={state.status === "permission" ? `Allow access to “${state.folder}”` : "No database folder"}>
        {state.status === "permission" ? (
          <Button size="sm" className="mt-2" onClick={() => void db.open(true)}>
            Allow access
          </Button>
        ) : state.status === "error" ? (
          <p className="text-destructive">{state.error}</p>
        ) : (
          <p>
            Choose the folder holding the database in{" "}
            <Link to="/settings" className="underline">
              Settings
            </Link>
            .
          </p>
        )}
      </EmptyState>
    </div>
  );
}

/** Banner when Power BI is not signed in. */
export function SignInBanner() {
  const { status } = useAuth();
  if (!status || status.signedIn) return null;
  return (
    <div className="flex items-center gap-2 border-b bg-muted/50 px-4 py-1.5 text-[12px]">
      <LogInIcon className="size-3.5 text-muted-foreground" /> Not signed in to Power BI.
      <Link to="/settings" className="underline underline-offset-2">
        Sign in
      </Link>
    </div>
  );
}
