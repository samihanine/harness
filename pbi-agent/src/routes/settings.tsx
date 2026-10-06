import { useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { CheckCircle2Icon, CopyIcon, ExternalLinkIcon, FolderIcon, LogOutIcon } from "lucide-react";
import { AGENT_URL } from "@/agent/server";
import { IconButton } from "@/components/icon-button";
import { Page } from "@/components/page";
import { Button } from "@/components/ui/button";
import { db, useDb } from "@/db/db";
import { logout, startLogin } from "@/pbi/auth";
import { useAuth } from "@/lib/use-auth";

export const Route = createFileRoute("/settings")({ component: Settings });

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section className="flex flex-col gap-3 rounded-xl border bg-card p-4 shadow-soft">
      <h2 className="text-[12px] font-medium text-muted-foreground uppercase">{title}</h2>
      {children}
    </section>
  );
}

function Settings() {
  const { status, error, refresh } = useAuth();
  const { state } = useDb();
  const [busy, setBusy] = useState(false);
  const [loginError, setLoginError] = useState("");

  const signIn = async () => {
    setBusy(true);
    setLoginError("");
    try {
      await startLogin();
      await refresh();
    } catch (e) {
      setLoginError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const chooseFolder = async () => {
    try {
      await db.choose(await window.showDirectoryPicker({ mode: "readwrite", id: "pbi-agent-db" }));
    } catch (e) {
      if ((e as Error).name !== "AbortError") alert((e as Error).message);
    }
  };

  return (
    <Page title="Settings">
      <Section title="Power BI">
        {error && <p className="text-[12px] text-destructive">{error}</p>}
        {status?.signedIn ? (
          <div className="flex items-center gap-2">
            <CheckCircle2Icon className="size-4 text-success" />
            <span>
              Signed in{status.account ? ` as ${status.account}` : ""}. Tokens are renewed automatically by the local API.
            </span>
            <IconButton label="Sign out" className="ml-auto" onClick={() => void logout().then(refresh)}>
              <LogOutIcon />
            </IconButton>
          </div>
        ) : status?.pending ? (
          <div className="flex flex-col gap-2">
            <p>
              Open{" "}
              <a href={status.pending.verificationUri} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 underline">
                {status.pending.verificationUri} <ExternalLinkIcon className="size-3" />
              </a>{" "}
              and enter this code:
            </p>
            <div className="flex items-center gap-2">
              <code className="rounded-lg border bg-muted px-3 py-1.5 font-mono text-[18px] tracking-[0.2em]">{status.pending.userCode}</code>
              <IconButton label="Copy" onClick={() => void navigator.clipboard.writeText(status.pending!.userCode)}>
                <CopyIcon />
              </IconButton>
            </div>
            <p className="shimmer w-fit text-[12px]">Waiting for the sign-in…</p>
            {status.pending.error && <p className="text-[12px] text-destructive">{status.pending.error}</p>}
            <Button variant="ghost" size="sm" className="w-fit" onClick={() => void signIn()}>
              New code
            </Button>
          </div>
        ) : (
          <div className="flex flex-col gap-2">
            <p className="text-muted-foreground">Sign in once with your work account (any organization): the local API keeps the session.</p>
            <Button size="sm" className="w-fit" disabled={busy} onClick={() => void signIn()}>
              Sign in to Power BI
            </Button>
            {loginError && <p className="text-[12px] text-destructive">{loginError}</p>}
          </div>
        )}
      </Section>

      <Section title="Database">
        <p className="text-muted-foreground">A local folder holding reports.xlsx, datasets.xlsx and guides.xlsx (created when missing).</p>
        <div className="flex items-center gap-2">
          <FolderIcon className="size-4 text-muted-foreground" />
          <span className="flex-1">
            {state.status === "ready" || state.status === "permission" ? state.folder : state.status === "error" ? state.error : "No folder"}
          </span>
          {state.status === "permission" && (
            <Button size="sm" variant="outline" onClick={() => void db.open(true)}>
              Allow access
            </Button>
          )}
          <Button size="sm" variant="outline" onClick={() => void chooseFolder()}>
            {state.status === "no-folder" ? "Choose a folder…" : "Change…"}
          </Button>
        </div>
      </Section>

      <Section title="Agent">
        <p className="text-muted-foreground">
          The agent panel is the ia-agent app at <code className="font-mono text-[12px]">{AGENT_URL}</code> (VITE_AGENT_URL to change it). Its AI key and
          settings are in its own settings.
        </p>
      </Section>
    </Page>
  );
}
