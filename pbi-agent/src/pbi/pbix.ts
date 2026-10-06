/**
 * .pbix files: reading the report definition (legacy Report/Layout or PBIR definition folder)
 * and writing a thin report connected live to a published semantic model.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type { Connection, FieldRef, PageContent, ReportContent, VisualContent } from "./types";

/* ---------------------------------- read ---------------------------------- */

const utf16 = (bytes: Uint8Array) => new TextDecoder("utf-16le").decode(bytes).replace(/^﻿/, "");
const json = (value: unknown) => {
  if (typeof value !== "string") return value ?? undefined;
  try {
    return JSON.parse(value);
  } catch {
    return undefined;
  }
};

const AGGREGATES = ["Sum", "Avg", "Count", "Min", "Max", "CountNonNull", "Median", "StandardDeviation", "Variance"];

type Expr = Record<string, any>;

/** Power BI query expression → "Table[Column]" / "[Measure]" / "Sum(Table[Column])". */
export function fieldOf(expr: Expr, aliases: Record<string, string> = {}): FieldRef {
  const entity = (e: Expr) => e?.SourceRef?.Entity ?? aliases[e?.SourceRef?.Source] ?? e?.SourceRef?.Source ?? "?";
  if (expr.Column) return `${entity(expr.Column.Expression)}[${expr.Column.Property}]`;
  if (expr.Measure) return `[${expr.Measure.Property}]`;
  if (expr.Aggregation) return `${AGGREGATES[expr.Aggregation.Function] ?? "Agg"}(${fieldOf(expr.Aggregation.Expression, aliases)})`;
  if (expr.HierarchyLevel) {
    const h = expr.HierarchyLevel.Expression?.Hierarchy;
    return `${entity(h?.Expression ?? h?.Expression?.PropertyVariationSource?.Expression)}[${h?.Hierarchy}.${expr.HierarchyLevel.Level}]`;
  }
  return JSON.stringify(expr).slice(0, 80);
}

/** Title of a visual from its container objects ("'Sales by region'" literal). */
const titleOf = (objects: any) => {
  const value = objects?.title?.[0]?.properties?.text?.expr?.Literal?.Value;
  return typeof value === "string" ? value.replace(/^'|'$/g, "") : undefined;
};

function legacyVisual(container: any): VisualContent | null {
  const config = json(container.config);
  const single = config?.singleVisual;
  if (!single) return null; // groups, decorative shapes without visual type
  const aliases: Record<string, string> = Object.fromEntries((single.prototypeQuery?.From ?? []).map((f: any) => [f.Name, f.Entity]));
  const refs: Record<string, FieldRef> = Object.fromEntries(
    (single.prototypeQuery?.Select ?? []).map((s: any) => [s.Name, fieldOf(s, aliases)]),
  );
  const fields: Record<string, FieldRef[]> = {};
  for (const [role, items] of Object.entries(single.projections ?? {}))
    fields[role] = (items as { queryRef: string }[]).map((p) => refs[p.queryRef] ?? p.queryRef);
  const layout = config.layouts?.[0]?.position ?? container;
  return {
    name: config.name,
    type: single.visualType,
    title: titleOf(single.vcObjects),
    x: Math.round(layout.x ?? 0),
    y: Math.round(layout.y ?? 0),
    z: layout.z,
    width: Math.round(layout.width ?? 0),
    height: Math.round(layout.height ?? 0),
    hidden: single.display?.mode === "hidden" || undefined,
    fields,
    filters: json(container.filters) ?? [],
    objects: single.objects,
    containerObjects: single.vcObjects,
    raw: config,
  };
}

function readLegacy(layout: any): Omit<ReportContent, "extractedAt" | "warnings"> {
  const pages: PageContent[] = (layout.sections ?? []).map((section: any) => ({
    name: section.name,
    displayName: section.displayName,
    width: section.width,
    height: section.height,
    hidden: json(section.config)?.visibility === 1 || undefined,
    filters: json(section.filters) ?? [],
    visuals: (section.visualContainers ?? []).map(legacyVisual).filter(Boolean),
  }));
  const config = json(layout.config);
  return { source: "pbix", pages, filters: json(layout.filters) ?? [], theme: config?.themeCollection };
}

function readPbir(files: Record<string, Uint8Array>): Omit<ReportContent, "extractedAt" | "warnings"> {
  const read = (path: string) => json(strFromU8(files[path]));
  const root = "Report/definition/";
  const pageDirs = [...new Set(Object.keys(files).filter((p) => p.startsWith(`${root}pages/`) && p.endsWith("/page.json")))];
  const order: string[] = (read(`${root}pages/pages.json`) as any)?.pageOrder ?? [];
  const pages: PageContent[] = pageDirs.map((path) => {
    const page = read(path) as any;
    const dir = path.replace(/page\.json$/, "");
    const visuals = Object.keys(files)
      .filter((p) => p.startsWith(`${dir}visuals/`) && p.endsWith("/visual.json"))
      .map((p) => read(p) as any)
      .filter((v) => v?.visual)
      .map((v): VisualContent => {
        const fields: Record<string, FieldRef[]> = {};
        for (const [role, state] of Object.entries(v.visual.query?.queryState ?? {}))
          fields[role] = ((state as any).projections ?? []).map((p: any) => fieldOf(p.field));
        return {
          name: v.name,
          type: v.visual.visualType,
          title: titleOf(v.visual.visualContainerObjects),
          x: Math.round(v.position?.x ?? 0),
          y: Math.round(v.position?.y ?? 0),
          z: v.position?.z,
          width: Math.round(v.position?.width ?? 0),
          height: Math.round(v.position?.height ?? 0),
          hidden: v.isHidden || undefined,
          fields,
          filters: v.filterConfig?.filters ?? [],
          objects: v.visual.objects,
          containerObjects: v.visual.visualContainerObjects,
          raw: v,
        };
      });
    return {
      name: page.name,
      displayName: page.displayName,
      width: page.width,
      height: page.height,
      hidden: page.visibility === "HiddenInViewMode" || undefined,
      filters: page.filterConfig?.filters ?? [],
      visuals,
    };
  });
  pages.sort((a, b) => order.indexOf(a.name) - order.indexOf(b.name));
  const report = read(`${root}report.json`) as any;
  return { source: "pbir", pages, filters: report?.filterConfig?.filters ?? [], theme: report?.themeCollection };
}

export type PbixInfo = { content: ReportContent; connection?: Connection; reportId?: string };

/** Reads a .pbix: report definition + live connection (dataset id) when there is one. */
export function readPbix(bytes: Uint8Array): PbixInfo {
  const files = unzipSync(bytes);
  const warnings: string[] = [];
  let base: Omit<ReportContent, "extractedAt" | "warnings">;
  if (files["Report/Layout"]) base = readLegacy(JSON.parse(utf16(files["Report/Layout"])));
  else if (Object.keys(files).some((p) => p.startsWith("Report/definition/"))) base = readPbir(files);
  else throw new Error("No report definition found in this file.");

  let connection: Connection | undefined;
  let reportId: string | undefined;
  if (files.Connections) {
    const data = json(strFromU8(files.Connections)) as any;
    const live = data?.Connections?.[0];
    const remote = data?.RemoteArtifacts?.[0];
    const datasetId =
      remote?.DatasetId ?? (/^[0-9a-f-]{36}$/i.test(live?.PbiModelDatabaseName ?? "") ? live.PbiModelDatabaseName : undefined) ??
      live?.ConnectionString?.match(/Initial Catalog=([0-9a-f-]{36})/i)?.[1];
    reportId = remote?.ReportId;
    if (datasetId) connection = { kind: "remote", datasetId: datasetId.toLowerCase() };
    else if (live?.ConnectionString)
      warnings.push(`Live connection without dataset id: ${live.ConnectionString.slice(0, 160)}. Add the dataset with its link.`);
  }
  if (!connection && files.DataModel) warnings.push("This file holds its own (imported) model: publish it, then add the published report.");
  return { content: { ...base, warnings, extractedAt: new Date().toISOString() }, connection, reportId };
}

/* ---------------------------------- write --------------------------------- */

const utf16Bytes = (text: string) => {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = code >> 8;
  }
  return out;
};

const CONTENT_TYPES = `<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="json" ContentType="" /><Override PartName="/Version" ContentType="" /><Override PartName="/DiagramLayout" ContentType="" /><Override PartName="/Report/Layout" ContentType="" /><Override PartName="/Settings" ContentType="" /><Override PartName="/Metadata" ContentType="" /><Override PartName="/Connections" ContentType="" /></Types>`;

/** Thin .pbix connected live to a published semantic model (no data inside). */
export function writePbix(layout: unknown, datasetId: string) {
  const connections = {
    Version: 1,
    Connections: [
      {
        Name: "EntityDataSource",
        ConnectionString: `Data Source=pbiazure://api.powerbi.com;Initial Catalog=${datasetId};Identity Provider="https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 929d0ec0-7a41-4b1e-bc7c-b754a28bddcc";Integrated Security=ClaimsToken`,
        ConnectionType: "pbiServiceLive",
        PbiServiceModelId: 0,
        PbiModelVirtualServerName: "sobe_wowvirtualserver",
        PbiModelDatabaseName: datasetId,
      },
    ],
    RemoteArtifacts: [{ DatasetId: datasetId, ReportId: crypto.randomUUID() }],
  };
  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    Version: utf16Bytes("1.28"),
    Connections: strToU8(JSON.stringify(connections)),
    "Report/Layout": utf16Bytes(JSON.stringify(layout)),
    Settings: utf16Bytes(JSON.stringify({ Version: 4, ReportSettings: {}, QueriesSettings: { TypeDetectionEnabled: true, RelationshipImportEnabled: true } })),
    Metadata: utf16Bytes(JSON.stringify({ Version: 5, AutoCreatedRelationships: [], FileDescription: "", CreatedFrom: "Cloud", CreatedFromRelease: "2024.01" })),
    DiagramLayout: utf16Bytes(JSON.stringify({ version: "1.1.0", diagrams: [] })),
  });
}

/** Measures referenced by a report ("[Name]" fields), to complete a model whose measures are not readable. */
export function measuresOf(content: ReportContent): string[] {
  const fields = content.pages.flatMap((p) => p.visuals.flatMap((v) => Object.values(v.fields).flat()));
  return fields.flatMap((f) => {
    const match = /^\[(.+)\]$/.exec(f);
    return match ? [match[1]] : [];
  });
}
