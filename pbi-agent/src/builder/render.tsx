/**
 * Instant preview of a builder visual: its data comes from the same query engine (DAX on the
 * dataset), drawn with ECharts / HTML in the spirit of the Power BI visual.
 */
import { useEffect, useMemo, useRef, useState } from "react";
import * as echarts from "echarts";
import { formatCell } from "@/components/result-table";
import type { Source } from "@/query/engine";
import type { Pivot, PivotResult } from "@/query/pivot";
import { layout, parseField, runPivot } from "@/query/pivot";
import type { BuilderVisual } from "./spec";
import { PALETTE, rolesOf } from "./spec";

/** Dimensions and values of a visual as a pivot on its dataset. */
export function visualPivot(visual: BuilderVisual, source: Source): { pivot: Pivot; dims: { role: string; label: string }[] } {
  const roles = rolesOf(visual.type);
  const dims: { role: string; label: string; ref: string }[] = [];
  const values: Pivot["values"] = [];
  for (const [role, fields] of Object.entries(visual.fields))
    for (const f of fields) {
      const { column, table } = parseField(f.ref);
      const isValue = !table || f.aggregate || roles[role]?.kind === "value";
      if (isValue) values.push({ label: `${role}:${column}`, source: source.id, field: f.ref, aggregate: table ? (f.aggregate ?? "sum") : undefined });
      else dims.push({ role, label: column, ref: f.ref });
    }
  const filters = visual.filters.filter((f) => f.values.length).map((f) => ({ source: source.id, field: f.ref, values: f.values }));
  const columnRole = visual.type === "pivotTable" ? "Columns" : "";
  return {
    pivot: {
      rows: dims.filter((d) => d.role !== columnRole).map((d) => ({ label: d.label, bindings: { [source.id]: d.ref } })),
      columns: dims.filter((d) => d.role === columnRole).map((d) => ({ label: d.label, bindings: { [source.id]: d.ref } })),
      values,
      filters,
    },
    dims,
  };
}

export function useVisualData(visual: BuilderVisual, source: Source | undefined) {
  const [state, setState] = useState<{ result?: PivotResult; error?: string; loading: boolean }>({ loading: false });
  const spec = useMemo(() => (source ? visualPivot(visual, source) : null), [visual, source]);
  const key = JSON.stringify(spec?.pivot);
  useEffect(() => {
    if (!spec || !source) return;
    if (!spec.pivot.values.length && !spec.pivot.rows.length) return setState({ loading: false });
    // A slicer / table of columns only: list distinct values with a count.
    const pivot = spec.pivot.values.length ? spec.pivot : { ...spec.pivot, values: [{ label: "n", source: source.id, field: spec.pivot.rows[0].bindings[source.id], aggregate: "count" as const }] };
    let cancelled = false;
    setState((s) => ({ ...s, loading: true }));
    runPivot(pivot, [source])
      .then((result) => {
        if (cancelled) return;
        const error = result.queries.find((q) => q.error)?.error;
        setState({ result: error ? undefined : result, error, loading: false });
      })
      .catch((e: Error) => !cancelled && setState({ error: e.message, loading: false }));
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key, source]);
  return { ...state, spec };
}

const css = (name: string) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();

export function VisualPreview({ visual, source }: { visual: BuilderVisual; source?: Source }) {
  const { result, error, loading, spec } = useVisualData(visual, source);
  const fmt = visual.format;
  const body = () => {
    if (error) return <p className="p-3 text-[12px] text-destructive">{error}</p>;
    if (!spec || (!spec.pivot.values.length && !spec.pivot.rows.length)) return <p className="p-3 text-[12px] text-muted-foreground">Add fields on the right.</p>;
    if (!result) return null;
    const rows = result.rows;
    switch (visual.type) {
      case "card":
        return (
          <div className="flex h-full flex-col items-center justify-center">
            <span className="text-[40px] font-semibold tracking-tight tabular-nums" style={{ color: fmt.colors[0] }}>
              {formatCell(rows[0]?.v0)}
            </span>
            <span className="text-[12px] text-muted-foreground">{spec.pivot.values[0]?.label.split(":")[1]}</span>
          </div>
        );
      case "slicer":
        return (
          <div className="flex flex-col gap-1 overflow-auto p-3 text-[12px]">
            {rows.map((r, i) => (
              <label key={i} className="flex items-center gap-2">
                <span className="size-3 rounded-sm border" /> {formatCell(r.d0)}
              </label>
            ))}
          </div>
        );
      case "tableEx":
      case "pivotTable":
        return <TablePreview visual={visual} pivot={spec.pivot} result={result} />;
      default:
        return <Chart visual={visual} pivot={spec.pivot} result={result} />;
    }
  };
  return (
    <div className="flex size-full flex-col overflow-hidden rounded-lg" style={{ background: fmt.background, fontSize: fmt.fontSize }}>
      {fmt.showTitle && visual.title && (
        <p className="shrink-0 px-3 pt-2 font-medium" style={{ fontSize: fmt.titleSize }}>
          {visual.title}
        </p>
      )}
      <div className={`relative min-h-0 flex-1 transition-opacity ${loading ? "opacity-50" : ""}`}>{body()}</div>
    </div>
  );
}

function TablePreview({ pivot, result }: { visual: BuilderVisual; pivot: Pivot; result: PivotResult }) {
  const grid = layout(pivot, result);
  return (
    <div className="size-full overflow-auto">
      <table className="w-full text-[12px]">
        <thead className="sticky top-0 bg-card">
          <tr className="border-b">
            {pivot.rows.map((d) => (
              <th key={d.label} className="px-2 py-1 text-left font-medium">
                {d.label}
              </th>
            ))}
            {grid.colKeys.flatMap((key) =>
              pivot.values.map((v) => (
                <th key={JSON.stringify(key) + v.label} className="px-2 py-1 text-right font-medium">
                  {[...key.map(formatCell), v.label.split(":")[1]].filter(Boolean).join(" · ")}
                </th>
              )),
            )}
          </tr>
        </thead>
        <tbody>
          {grid.rows.slice(0, 200).map((row) => (
            <tr key={JSON.stringify(row.labels)} className="border-b border-border/50">
              {row.labels.map((l, i) => (
                <td key={i} className="px-2 py-1">
                  {formatCell(l)}
                </td>
              ))}
              {grid.colKeys.flatMap((key) =>
                pivot.values.map((_, vi) => (
                  <td key={JSON.stringify(key) + vi} className="px-2 py-1 text-right tabular-nums">
                    {formatCell(grid.cell(row.cells, key, vi))}
                  </td>
                )),
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Chart({ visual, pivot, result }: { visual: BuilderVisual; pivot: Pivot; result: PivotResult }) {
  const element = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const chart = echarts.init(element.current!, undefined, { renderer: "svg" });
    chart.setOption(chartOption(visual, pivot, result), true);
    const observer = new ResizeObserver(() => chart.resize());
    observer.observe(element.current!);
    return () => {
      observer.disconnect();
      chart.dispose();
    };
  }, [visual, pivot, result]);
  return <div ref={element} className="size-full" />;
}

function chartOption(visual: BuilderVisual, pivot: Pivot, result: PivotResult): echarts.EChartsOption {
  const fmt = visual.format;
  const text = css("--muted-foreground") || "#888";
  const grid = css("--border") || "#eee";
  const colors = fmt.colors.length ? [...fmt.colors, ...PALETTE] : PALETTE;
  const rows = result.rows;
  const hasSeries = (visual.fields.Series?.length ?? 0) > 0;
  const categories = [...new Set(rows.map((r) => formatCell(r.d0)))];
  const valueName = (i: number) => pivot.values[i]?.label.split(":")[1] ?? "";
  const label = { show: fmt.showDataLabels, color: text, fontSize: fmt.fontSize ?? 11 };
  const legend = {
    show: fmt.showLegend,
    type: "scroll" as const,
    pageTextStyle: { color: text },
    textStyle: { color: text, fontSize: 11 },
    ...(fmt.legendPosition === "Bottom" ? { bottom: 0 } : fmt.legendPosition === "Left" ? { left: 0, orient: "vertical" as const } : fmt.legendPosition === "Right" ? { right: 0, orient: "vertical" as const } : { top: 0 }),
  };
  const base = { color: colors, backgroundColor: "transparent", tooltip: { trigger: "item" as const }, legend, animationDuration: 300 };

  if (visual.type === "pieChart" || visual.type === "donutChart")
    return {
      ...base,
      series: [{ type: "pie", radius: visual.type === "donutChart" ? ["45%", "70%"] : "70%", label: { ...label, show: fmt.showDataLabels }, data: rows.map((r) => ({ name: formatCell(r.d0), value: Number(r.v0) || 0 })) }],
    };
  if (visual.type === "gauge") {
    const value = Number(rows[0]?.v0) || 0;
    const max = Number(rows[0]?.v1) || value * 1.5 || 100;
    return { ...base, series: [{ type: "gauge", max, progress: { show: true }, detail: { valueAnimation: true, color: text, fontSize: 20 }, data: [{ value, name: valueName(0) }], axisLabel: { color: text } }] };
  }
  if (visual.type === "scatterChart")
    return {
      ...base,
      grid: { left: 40, right: 16, top: 24, bottom: 32, containLabel: true },
      xAxis: { type: "value", name: fmt.showAxisTitles ? valueName(0) : "", axisLabel: { color: text }, splitLine: { lineStyle: { color: grid } } },
      yAxis: { type: "value", name: fmt.showAxisTitles ? valueName(1) : "", axisLabel: { color: text }, splitLine: { lineStyle: { color: grid } } },
      series: [{ type: "scatter", label: { ...label, formatter: "{b}" }, data: rows.map((r) => ({ name: formatCell(r.d0), value: [Number(r.v0), Number(r.v1)] })) }],
    };

  const horizontal = visual.type === "clusteredBarChart" || visual.type === "barChart";
  const stacked = visual.type === "columnChart" || visual.type === "barChart" || (visual.type === "areaChart" && hasSeries);
  const seriesType = visual.type === "lineChart" || visual.type === "areaChart" ? "line" : "bar";
  const series = hasSeries
    ? [...new Set(rows.map((r) => formatCell(r.d1)))].map((name) => ({
        name,
        data: categories.map((c) => rows.find((r) => formatCell(r.d0) === c && formatCell(r.d1) === name)?.v0 ?? null),
      }))
    : pivot.values.map((_, i) => ({ name: valueName(i), data: categories.map((c) => rows.find((r) => formatCell(r.d0) === c)?.[`v${i}`] ?? null) }));
  const categoryAxis = { type: "category" as const, data: categories, name: fmt.showAxisTitles ? pivot.rows[0]?.label : "", nameLocation: "middle" as const, nameGap: 28, axisLabel: { color: text }, axisLine: { lineStyle: { color: grid } } };
  const valueAxis = {
    type: "value" as const,
    name: fmt.showAxisTitles ? (hasSeries ? valueName(0) : series.map((s) => s.name).join(", ")) : "",
    nameLocation: "middle" as const,
    nameGap: horizontal ? 28 : 48,
    nameRotate: horizontal ? 0 : 90,
    axisLabel: { color: text },
    splitLine: { lineStyle: { color: grid } },
  };
  return {
    ...base,
    tooltip: { trigger: "axis" },
    grid: { left: fmt.showAxisTitles && !horizontal ? 36 : 16, right: 16, top: fmt.showLegend && fmt.legendPosition === "Top" ? 36 : 16, bottom: fmt.showAxisTitles ? 36 : 16, containLabel: true },
    xAxis: horizontal ? valueAxis : categoryAxis,
    yAxis: horizontal ? categoryAxis : valueAxis,
    series: series.map((s) => ({
      ...s,
      type: seriesType,
      stack: stacked ? "total" : undefined,
      areaStyle: visual.type === "areaChart" ? { opacity: 0.25 } : undefined,
      smooth: false,
      label,
      barMaxWidth: 40,
    })) as echarts.SeriesOption[],
  };
}
