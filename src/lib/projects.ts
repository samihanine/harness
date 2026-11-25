/** What the server keeps of each project's Confluence pages (read once, refreshed on demand), shared by the screens. */
import { useEffect, useState } from "react";
import { get, post } from "./api";

export type ProjectStatus = { id: string; pages: number; at?: string; url?: string; error?: string };

let loaded: Promise<ProjectStatus[]> | undefined;
const listeners = new Set<(s: ProjectStatus[]) => void>();
const publish = (p: Promise<ProjectStatus[]>) => ((loaded = p), p.then((s) => listeners.forEach((l) => l(s)), () => (loaded = undefined)), p);

export function useProjectsStatus(enabled = true) {
  const [status, setStatus] = useState<ProjectStatus[]>();
  useEffect(() => {
    if (!enabled) return;
    listeners.add(setStatus);
    (loaded ?? publish(get<ProjectStatus[]>("/projects/status"))).then(setStatus, () => undefined);
    return () => void listeners.delete(setStatus);
  }, [enabled]);
  return status;
}

/** Reads a project's pages from Confluence again (the next conversations get them). */
export const refreshProject = (id: string) => publish(post<ProjectStatus[]>("/projects/refresh", { id }));
