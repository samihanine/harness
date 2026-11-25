/**
 * The Builder's editor mode, browser side: the report is embedded in the Power BI editor
 * (powerbi-report-authoring) and these actions change it live — pages, visuals, fields, formatting,
 * positions, theme — then save it. For reports whose definition the API cannot write (PBIR-Legacy format,
 * or refused): what the v1 "editor mode" did. Fields arrive as checked targets (the server reads the model).
 */
import type * as PbiClient from "powerbi-client";
// The authoring SDK adds the editing methods (createVisual, addDataField, setProperty…) to the client.
import type {} from "powerbi-report-authoring";
import type { BrowserActions } from "@/lib/chat";
import { powerbi, type Report } from "@/features/powerbi/embed";
import { describeError } from "./report-actions";

type Visual = PbiClient.VisualDescriptor;
type Target = { table: string; column?: string; measure?: string; aggregationFunction?: string };
const PROPERTY = "http://powerbi.com/product/schema#property";
const ROLE_ALIASES: Record<string, string> = { SmallMultiples: "Rows" };

const fieldName = (t?: Record<string, any> | null) => (!t ? "?" : t.measure ? `[${t.measure}]` : `${t.aggregationFunction ? `${t.aggregationFunction}(` : ""}'${t.table}'[${t.column ?? t.hierarchyLevel ?? t.hierarchy}]${t.aggregationFunction ? ")" : ""}`);

/** Deep merge of plain objects (lists replaced). */
const merge = (a: any, b: any): any => {
  if (!a || typeof a !== "object" || Array.isArray(a) || !b || typeof b !== "object" || Array.isArray(b)) return b;
  const out = { ...a };
  for (const [k, v] of Object.entries(b)) out[k] = merge(a[k], v);
  return out;
};

export function editorActions(report: () => Report | undefined, ready: () => Promise<void>): BrowserActions {
  const r = async () => {
    await ready();
    const x = report();
    if (!x) throw new Error("The report is not open in the editor.");
    return x;
  };
  const pageOf = async (page?: string) => {
    const pages = await (await r()).getPages();
    const p = page ? pages.find((x) => x.name === page || x.displayName.toLowerCase() === page.toLowerCase()) : pages.find((x) => x.isActive);
    if (!p) throw new Error(`No page "${page}". Pages: ${pages.map((x) => `"${x.displayName}" (${x.name})`).join(", ")}.`);
    return p;
  };
  const visualOf = async (visual: string, page?: string) => {
    const visuals = await (await pageOf(page)).getVisuals();
    const v = visuals.find((x) => x.name === visual) ?? visuals.find((x) => (x.title ?? "").toLowerCase() === visual.toLowerCase());
    if (!v) throw new Error(`No visual "${visual}". Visuals: ${visuals.map((x) => `${x.type} "${x.title ?? ""}" (${x.name})`).join(", ")}.`);
    return v;
  };
  const setFields = async (v: Visual, fields: Record<string, Target[]>) => {
    for (const [given, list] of Object.entries(fields)) {
      const role = ROLE_ALIASES[given] ?? given;
      const current = await v.getDataFields(role).catch(() => []);
      for (let i = current.length - 1; i >= 0; i--) await v.removeDataField(role, i);
      for (const t of list) await v.addDataField(role, t as never);
    }
  };
  /** Each format property on its own: what the editor refuses is listed, the rest applied. */
  const format = async (v: Visual, given: Record<string, Record<string, unknown>>) => {
    const refused: string[] = [];
    for (const [objectName, props] of Object.entries(given ?? {}))
      for (const [propertyName, value] of Object.entries(props ?? {}))
        await v.setProperty({ objectName, propertyName }, { schema: PROPERTY, value } as never).catch((e: unknown) => refused.push(`${objectName}.${propertyName} (${describeError(e)})`));
    return refused;
  };
  const fieldsOf = async (v: Visual) => {
    if (typeof v.getCapabilities !== "function") return {};
    const caps = await v.getCapabilities().catch(() => null);
    const out: Record<string, string[]> = {};
    for (const role of caps?.dataRoles ?? []) {
      const f = await v.getDataFields(role.name).catch(() => []);
      if (f.length) out[role.name] = f.map((x) => fieldName(x as never));
    }
    return out;
  };
  return {
    "editor.overview": async ({ page }: { page?: string }) => {
      const pages = await (await r()).getPages();
      const shown = page ? [await pageOf(page)] : pages;
      return Promise.all(
        shown.map(async (p) => ({
          name: p.name,
          displayName: p.displayName,
          active: p.isActive,
          size: p.defaultSize,
          visuals: await Promise.all((await p.getVisuals()).map(async (v) => ({ name: v.name, type: v.type, title: v.title, x: Math.round(v.layout.x ?? 0), y: Math.round(v.layout.y ?? 0), width: Math.round(v.layout.width ?? 0), height: Math.round(v.layout.height ?? 0), hidden: v.layout.displayState?.mode === 1 || undefined, fields: await fieldsOf(v) }))),
        })),
      );
    },
    "editor.addPage": async ({ displayName }: { displayName: string }) => {
      const p = await (await r()).addPage(displayName);
      await p.setActive();
      return { name: p.name };
    },
    "editor.renamePage": async ({ page, displayName }: { page: string; displayName: string }) => (await (await pageOf(page)).setDisplayName(displayName), { ok: true }),
    "editor.deletePage": async ({ page }: { page: string }) => {
      const p = await pageOf(page);
      await (await r()).deletePage(p.name);
      return { ok: true };
    },
    "editor.addVisual": async ({ page, type, layout, fields, title, format: given }: { page?: string; type: string; layout: { x: number; y: number; width: number; height: number }; fields?: Record<string, Target[]>; title?: string; format?: Record<string, Record<string, unknown>> }) => {
      const p = await pageOf(page);
      // Without an explicit display state the editor makes the visual hidden.
      const { visual } = await p.createVisual(type, { ...layout, displayState: { mode: 0 } } as never);
      await p.setVisualDisplayState(visual.name, 0 as never).catch(() => undefined);
      const v = (await p.getVisuals()).find((x) => x.name === visual.name)!;
      try {
        if (fields) await setFields(v, fields);
      } catch (e) {
        // A visual whose fields are refused is removed: nothing half made stays on the page.
        await p.deleteVisual(visual.name).catch(() => undefined);
        throw new Error(`Fields refused for a ${type}: ${describeError(e)}`);
      }
      const refused = await format(v, { ...(given ?? {}), ...(title ? { title: { visible: true, titleText: title, ...(given?.title ?? {}) } } : {}) });
      return { name: visual.name, refused };
    },
    "editor.updateVisual": async ({ page, visual, type, fields }: { page?: string; visual: string; type?: string; fields?: Record<string, Target[]> }) => {
      const v = await visualOf(visual, page);
      if (type && type !== v.type) await v.changeType(type);
      if (fields) await setFields(v, fields);
      const now = await visualOf(v.name, page);
      return { type: now.type, fields: await fieldsOf(now) };
    },
    "editor.format": async ({ page, visuals, format: given }: { page?: string; visuals: string[]; format: Record<string, Record<string, unknown>> }) => {
      const refused: string[] = [];
      for (const name of visuals) refused.push(...(await format(await visualOf(name, page), given)));
      return { refused: [...new Set(refused)] };
    },
    "editor.move": async ({ page, boxes }: { page?: string; boxes: { visual: string; x?: number; y?: number; width?: number; height?: number }[] }) => {
      for (const b of boxes) {
        const v = await visualOf(b.visual, page);
        if (b.x !== undefined || b.y !== undefined) await v.moveVisual(b.x ?? v.layout.x ?? 0, b.y ?? v.layout.y ?? 0);
        if (b.width !== undefined || b.height !== undefined) await v.resizeVisual(b.width ?? v.layout.width ?? 300, b.height ?? v.layout.height ?? 200);
      }
      return { ok: true };
    },
    "editor.deleteVisual": async ({ page, visual }: { page?: string; visual: string }) => {
      const p = await pageOf(page);
      const v = await visualOf(visual, page);
      await p.deleteVisual(v.name);
      return { ok: true };
    },
    "editor.theme": async ({ theme }: { theme: Record<string, unknown> }) => {
      const x = await r();
      const current = ((await x.getTheme().catch(() => undefined)) as { themeJson?: Record<string, unknown> } | undefined)?.themeJson ?? { name: "Hub theme" };
      await x.applyTheme({ themeJson: merge(current, theme) } as never);
      return { ok: true };
    },
    "editor.save": async () => {
      await (await r()).save();
      return { ok: true };
    },
    "editor.pbi": async () => (await powerbi(), { ok: true }),
  };
}
