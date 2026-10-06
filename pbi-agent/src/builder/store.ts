/** Builder reports, kept in this browser (IndexedDB). */
import { useSyncExternalStore } from "react";
import { createStore, get, set } from "idb-keyval";
import type { BuilderReport } from "./spec";

const store = createStore("pbi-agent-builder", "reports");
let reports: BuilderReport[] = [];
const listeners = new Set<() => void>();
const emit = () => listeners.forEach((l) => l());

void get<BuilderReport[]>("all", store).then((saved) => {
  reports = saved ?? [];
  emit();
});

export const builderReports = {
  all: () => reports,
  get: (id: string) => reports.find((r) => r.id === id),
  put(report: BuilderReport) {
    reports = reports.some((r) => r.id === report.id) ? reports.map((r) => (r.id === report.id ? report : r)) : [...reports, report];
    emit();
    void set("all", reports, store);
  },
  remove(id: string) {
    reports = reports.filter((r) => r.id !== id);
    emit();
    void set("all", reports, store);
  },
  use: () =>
    useSyncExternalStore(
      (l) => {
        listeners.add(l);
        return () => listeners.delete(l);
      },
      () => reports,
    ),
};
