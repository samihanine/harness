/** UI of the storage layer: Microsoft sign-in, local / SharePoint switch, SharePoint links. */
import { useCallback, useEffect, useState } from "react";
import { CheckCircle2Icon, CloudIcon, CopyIcon, ExternalLinkIcon, HardDriveIcon, LinkIcon, LogOutIcon, XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { cn } from "@/lib/utils";
import type { AuthStatus } from "@/storage/graph";
import { auth, resolveLink } from "@/storage/graph";
import type { DriveRef } from "@/storage/types";
import { IconButton } from "./icon-button";

/** Microsoft sign-in state (polled while a code is pending). */
export function useMicrosoft() {
  const [status, setStatus] = useState<AuthStatus | null>(null);
  const [error, setError] = useState("");
  const refresh = useCallback(
    () =>
      auth
        .status()
        .then((s) => (setStatus(s), setError("")))
        .catch((e: Error) => setError(e.message)),
    [],
  );
  useEffect(() => {
    void refresh();
  }, [refresh]);
  useEffect(() => {
    if (!status?.pending) return;
    const timer = setInterval(() => void refresh(), 3000);
    return () => clearInterval(timer);
  }, [status?.pending, refresh]);
  return { status, error, refresh };
}

/** Sign-in block: account when signed in, device code while signing in, else a button. */
export function MicrosoftSignIn({ onSignedIn }: { onSignedIn?: () => void }) {
  const { status, error, refresh } = useMicrosoft();
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    if (status?.signedIn) onSignedIn?.();
  }, [status?.signedIn]); // eslint-disable-line react-hooks/exhaustive-deps

  if (error) return <p className="text-[12px] text-destructive">{error}</p>;
  if (!status) return <p className="shimmer text-[12px]">Checking the Microsoft account…</p>;
  if (status.signedIn)
    return (
      <div className="flex items-center gap-2 text-[12px]">
        <CheckCircle2Icon className="size-3.5 text-success" />
        <span className="truncate">{status.account || "Signed in to Microsoft"}</span>
        <IconButton label="Sign out" size="icon-xs" className="ml-auto" onClick={() => void auth.logout().then(refresh)}>
          <LogOutIcon />
        </IconButton>
      </div>
    );
  const start = async () => {
    setBusy(true);
    await auth.login().catch(() => undefined);
    await refresh();
    setBusy(false);
  };
  if (status.pending)
    return (
      <div className="flex flex-col gap-1.5 rounded-lg border bg-muted/30 p-2.5 text-[12px]">
        <p>
          Open{" "}
          <a href={status.pending.verificationUri} target="_blank" rel="noreferrer" className="inline-flex items-center gap-0.5 underline">
            {status.pending.verificationUri.replace("https://", "")} <ExternalLinkIcon className="size-3" />
          </a>{" "}
          and enter:
        </p>
        <div className="flex items-center gap-1.5">
          <code className="rounded-md border bg-background px-2 py-1 font-mono text-[15px] tracking-[0.15em]">{status.pending.userCode}</code>
          <IconButton label="Copy" size="icon-xs" onClick={() => void navigator.clipboard.writeText(status.pending!.userCode)}>
            <CopyIcon />
          </IconButton>
          <span className="shimmer ml-auto">Waiting…</span>
        </div>
        {status.pending.error && <p className="text-destructive">{status.pending.error}</p>}
      </div>
    );
  return (
    <div className="flex items-center gap-2">
      <Button size="sm" variant="outline" disabled={busy} onClick={() => void start()}>
        Sign in with Microsoft
      </Button>
      <span className="text-[11px] text-muted-foreground">Work account, any organization.</span>
    </div>
  );
}

export type StorageKind = "local" | "sharepoint";

export function StorageTabs({ value, onChange }: { value: StorageKind; onChange: (kind: StorageKind) => void }) {
  return (
    <div className="flex rounded-lg bg-muted p-0.5">
      {(
        [
          { kind: "local", label: "This computer", icon: HardDriveIcon },
          { kind: "sharepoint", label: "SharePoint / OneDrive", icon: CloudIcon },
        ] as const
      ).map(({ kind, label, icon: Icon }) => (
        <button
          key={kind}
          type="button"
          onClick={() => onChange(kind)}
          className={cn(
            "flex flex-1 items-center justify-center gap-1.5 rounded-md px-2 py-1 text-[12px] text-muted-foreground transition-all",
            value === kind && "bg-background text-foreground shadow-soft",
          )}
        >
          <Icon className="size-3.5" /> {label}
        </button>
      ))}
    </div>
  );
}

/** Paste a SharePoint / OneDrive link (file or folder), resolved to a drive item. */
export function DriveLinkField({
  kind,
  value,
  onChange,
  placeholder,
}: {
  kind: "file" | "folder";
  value?: DriveRef;
  onChange: (ref: DriveRef | undefined) => void;
  placeholder: string;
}) {
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: boolean; error?: string }>({});
  if (value)
    return (
      <div className="flex h-8 items-center gap-2 rounded-lg border bg-muted/40 px-2.5 [&_svg]:size-3.5 [&_svg]:text-muted-foreground">
        <CloudIcon />
        <a href={value.webUrl} target="_blank" rel="noreferrer" className="flex-1 truncate hover:underline">
          {value.name}
        </a>
        <button type="button" aria-label="Clear" onClick={() => onChange(undefined)}>
          <XIcon />
        </button>
      </div>
    );
  const resolve = async () => {
    if (!link.trim()) return;
    setState({ busy: true });
    try {
      const item = await resolveLink(link);
      if ((kind === "folder") !== item.isFolder) throw new Error(kind === "folder" ? "This link is not a folder." : "This link is not a file.");
      if (kind === "file" && !/\.xlsx$/i.test(item.name)) throw new Error("Choose an .xlsx file.");
      const { isFolder: _, ...ref } = item;
      onChange(ref);
      setLink("");
      setState({});
    } catch (e) {
      setState({ error: e instanceof Error ? e.message : String(e) });
    }
  };
  return (
    <div className="flex flex-col gap-1">
      <div className="flex gap-1.5">
        <Input value={link} onChange={(e) => setLink(e.target.value)} onKeyDown={(e) => e.key === "Enter" && void resolve()} placeholder={placeholder} />
        <Button size="sm" variant="outline" disabled={!link.trim() || state.busy} onClick={() => void resolve()}>
          <LinkIcon /> {state.busy ? "…" : "Use"}
        </Button>
      </div>
      {state.error && <p className="text-[11px] text-destructive">{state.error}</p>}
    </div>
  );
}
