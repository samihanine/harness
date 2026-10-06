export function downloadJson(value: unknown, name: string) {
  const blob = new Blob([JSON.stringify(value, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${name.replace(/[^\w-]+/g, "-").replace(/^-|-$/g, "").toLowerCase() || "conversation"}.json`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export const duration = (start: number, end?: number) => {
  const ms = (end ?? Date.now()) - start;
  return ms < 1000 ? `${ms}ms` : ms < 60_000 ? `${(ms / 1000).toFixed(1)}s` : `${Math.round(ms / 60_000)}m`;
};
