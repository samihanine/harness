/**
 * An orange banner on top of every screen when something keeps the app from working: not signed in
 * to Microsoft, or no AI key / a key the AI API refuses. Closed for this visit only: it comes back on
 * reload while the problem lasts.
 */
import { useEffect, useState } from "react";
import { Link, useLocation } from "@tanstack/react-router";
import { TriangleAlertIcon, XIcon } from "lucide-react";
import { aiProblem } from "@/services/health";

export function HealthBanner({ signedIn }: { signedIn?: boolean }) {
  const path = useLocation({ select: (l) => l.pathname });
  const [ai, setAi] = useState<string>();
  const [checked, setChecked] = useState(false);
  const [closed, setClosed] = useState(false);
  // Checked at start, then again on each screen change while there is a problem (e.g. after Settings).
  useEffect(() => {
    if (checked && !ai) return;
    void aiProblem().then((p) => (setAi(p), setChecked(true)));
  }, [path]); // eslint-disable-line react-hooks/exhaustive-deps
  const problems = [
    signedIn === false && { text: "You are not signed in to Microsoft: reports, models and files cannot be read.", tab: "powerbi", action: "Sign in" },
    ai && { text: ai, tab: "ai", action: "Set the AI key" },
  ].filter((p): p is { text: string; tab: string; action: string } => !!p);
  if (closed || !problems.length) return null;
  return (
    <div role="alert" className="flex shrink-0 items-start gap-3 border-b border-warning/30 bg-warning/15 px-5 py-2 text-[13px]">
      <TriangleAlertIcon className="mt-0.5 size-4 shrink-0 text-warning" aria-hidden />
      <div className="flex min-w-0 flex-1 flex-col gap-0.5">
        {problems.map((p) => (
          <span key={p.tab}>
            {p.text}{" "}
            <Link to="/settings" search={{ tab: p.tab } as never} className="font-medium underline underline-offset-2">
              {p.action}
            </Link>
          </span>
        ))}
      </div>
      <button type="button" className="icon-btn size-6 shrink-0" aria-label="Close" onClick={() => setClosed(true)}>
        <XIcon className="size-3.5" />
      </button>
    </div>
  );
}
