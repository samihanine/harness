/**
 * A report in the old format (PBIR-Legacy) shown as PBIR and written back: unchanged, report.json comes back
 * the same; one visual moved, only its container changes.  bun scripts/checks/legacy.ts <reportId> [groupId]
 */
import { anyReportDefinition, partText } from "@server/pbi/definitions";
import { legacyToPbir, pbirToLegacy } from "@server/pbi/legacy";
import { textPart } from "@server/pbi/definitions";
const ref = { id: process.argv[2] ?? "e799444e-aaa6-4e09-a1b2-7b8d1fa6d898", groupId: process.argv[3] };
const parts = await anyReportDefinition(ref);
const original = JSON.parse(partText(parts.find((p) => p.path === "report.json")!));
const { parts: files, legacy } = legacyToPbir(parts);
console.log("PBIR files:", files.map((f) => f.path).join("\n  "));
const same = JSON.parse(partText(pbirToLegacy(files, legacy, new Set()).find((p) => p.path === "report.json")!));
const norm = (r: any) => JSON.stringify({ ...r, config: JSON.parse(r.config), sections: r.sections.map((s: any) => ({ ...s, config: JSON.parse(s.config), visualContainers: s.visualContainers.map((c: any) => ({ ...c, config: JSON.parse(c.config) })) })) });
console.log("unchanged round trip identical:", norm(same) === norm(original));
const visual = files.find((f) => f.path.endsWith("/visual.json"))!;
const v = JSON.parse(partText(visual));
v.position.x += 10;
const moved = files.map((f) => (f === visual ? textPart(f.path, JSON.stringify(v)) : f));
const after = JSON.parse(partText(pbirToLegacy(moved, legacy, new Set([visual.path])).find((p) => p.path === "report.json")!));
const diff = after.sections.flatMap((s: any, i: number) => s.visualContainers.filter((c: any, j: number) => JSON.stringify(c) !== JSON.stringify(original.sections[i].visualContainers[j])).map((c: any) => JSON.parse(c.config).name));
console.log("changed containers after moving", v.name, ":", diff);
if (norm(same) !== norm(original)) {
  const a = JSON.parse(norm(same)), b = JSON.parse(norm(original));
  for (const k of Object.keys({ ...a, ...b })) if (JSON.stringify(a[k]) !== JSON.stringify(b[k])) console.log("differs:", k, JSON.stringify(a[k]).slice(0, 300), "\n   was:", JSON.stringify(b[k]).slice(0, 300));
}
