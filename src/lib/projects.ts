/**
 * The projects (sub-folders of the Confluence projects folder) with what the server keeps of each (pages read
 * once, refreshed on demand), shared by the screens.
 */
import { useEffect, useState } from "react";
import type { Project } from "@shared/types";
import { get } from "./api";

let loaded: Promise<Project[]> | undefined;
const listeners = new Set<(s: Project[]) => void>();
const publish = (p: Promise<Project[]>) => ((loaded = p), p.then((s) => listeners.forEach((l) => l(s)), () => (loaded = undefined)), p);

export function useProjectList() {
  const [projects, setProjects] = useState<Project[]>();
  const [error, setError] = useState<string>();
  useEffect(() => {
    listeners.add(setProjects);
    (loaded ?? publish(get<Project[]>("/projects"))).then(setProjects, (e) => setError(e instanceof Error ? e.message : String(e)));
    return () => void listeners.delete(setProjects);
  }, []);
  return { projects, error };
}

/** Reads a project's pages from Confluence again ("*": the list of projects). */
export const refreshProject = (id: string) => publish(get<Project[]>(`/projects?refresh=${encodeURIComponent(id)}`));
