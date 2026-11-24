/** The blocks of an analysis as the user sees them: text, key figures, tables, charts, pivots. */
import { Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, Legend, Line, LineChart, Pie, PieChart, ResponsiveContainer, Scatter, ScatterChart, Tooltip, XAxis, YAxis } from "recharts";
import { pivotRows, type AnalysisBlock, type Rows } from "@shared/analysis";
import { Markdown } from "@/components/markdown";

const COLORS = ["var(--chart-1)", "var(--chart-2)", "var(--chart-3)", "var(--chart-4)", "var(--chart-5)"];

export function formatValue(v: unknown, format?: string) {
  if (typeof v !== "number") return v === null || v === undefined ? "—" : String(v);
  if (format === "percent") return `${(v * 100).toLocaleString(undefined, { maximumFractionDigits: 1 })} %`;
  if (format?.startsWith("currency:")) return v.toLocaleString(undefined, { style: "currency", currency: format.slice(9), maximumFractionDigits: 0 });
  if (format === "compact" || Math.abs(v) >= 1e6) return Intl.NumberFormat(undefined, { notation: "compact", maximumFractionDigits: 2 }).format(v);
  return v.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function DataTable({ rows, columns, max = 200 }: { rows: Rows; columns?: string[]; max?: number }) {
  const cols = columns ?? Object.keys(rows[0] ?? {});
  return (
    <div className="max-h-[420px] overflow-auto rounded-xl border">
      <table className="w-full text-[12.5px]">
        <thead className="sticky top-0 bg-muted">
          <tr>{cols.map((c) => <th key={c} className="th whitespace-nowrap">{c}</th>)}</tr>
        </thead>
        <tbody>
          {rows.slice(0, max).map((r, i) => (
            <tr key={i} className="border-t">
              {cols.map((c) => (
                <td key={c} className={typeof r[c] === "number" ? "px-3 py-1.5 text-right tabular-nums" : "px-3 py-1.5"}>{formatValue(r[c])}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {rows.length > max && <div className="border-t px-3 py-1.5 text-[12px] text-muted-foreground">{rows.length - max} more rows (in the Excel export)</div>}
    </div>
  );
}

function Chart({ b }: { b: Extract<AnalysisBlock, { type: "chart" }> }) {
  const rows = b.rows ?? [];
  const { kind, x, y, stacked } = b.chart;
  const axis = { stroke: "var(--muted-foreground)", fontSize: 11, tickLine: false };
  const series = y.map((s, i) => ({ key: s, color: COLORS[i % COLORS.length] }));
  return (
    <div className="h-[320px]">
      <ResponsiveContainer>
        {kind === "pie" ? (
          <PieChart>
            <Pie data={rows} dataKey={y[0]} nameKey={x} innerRadius="45%" outerRadius="80%" paddingAngle={1}>
              {rows.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
            </Pie>
            <Tooltip formatter={(v) => formatValue(v)} />
            <Legend />
          </PieChart>
        ) : kind === "line" ? (
          <LineChart data={rows}>
            <CartesianGrid vertical={false} stroke="var(--border)" />
            <XAxis dataKey={x} {...axis} />
            <YAxis {...axis} tickFormatter={(v) => formatValue(v, "compact") as string} />
            <Tooltip formatter={(v) => formatValue(v)} />
            {series.length > 1 && <Legend />}
            {series.map((s) => <Line key={s.key} dataKey={s.key} stroke={s.color} strokeWidth={2} dot={false} />)}
          </LineChart>
        ) : kind === "area" ? (
          <AreaChart data={rows}>
            <CartesianGrid vertical={false} stroke="var(--border)" />
            <XAxis dataKey={x} {...axis} />
            <YAxis {...axis} tickFormatter={(v) => formatValue(v, "compact") as string} />
            <Tooltip formatter={(v) => formatValue(v)} />
            {series.map((s) => <Area key={s.key} dataKey={s.key} stroke={s.color} fill={s.color} fillOpacity={0.2} stackId={stacked ? "a" : undefined} />)}
          </AreaChart>
        ) : kind === "scatter" ? (
          <ScatterChart>
            <CartesianGrid stroke="var(--border)" />
            <XAxis dataKey={x} {...axis} type="number" />
            <YAxis dataKey={y[0]} {...axis} type="number" />
            <Tooltip formatter={(v) => formatValue(v)} />
            <Scatter data={rows} fill={COLORS[0]} />
          </ScatterChart>
        ) : (
          <BarChart data={rows} layout={kind === "bar" ? "vertical" : "horizontal"}>
            <CartesianGrid vertical={kind !== "bar"} horizontal={kind === "bar" ? false : true} stroke="var(--border)" />
            {kind === "bar" ? (
              <>
                <XAxis type="number" {...axis} tickFormatter={(v) => formatValue(v, "compact") as string} />
                <YAxis type="category" dataKey={x} {...axis} width={120} />
              </>
            ) : (
              <>
                <XAxis dataKey={x} {...axis} />
                <YAxis {...axis} tickFormatter={(v) => formatValue(v, "compact") as string} />
              </>
            )}
            <Tooltip formatter={(v) => formatValue(v)} />
            {series.length > 1 && <Legend />}
            {series.map((s) => <Bar key={s.key} dataKey={s.key} fill={s.color} radius={3} stackId={stacked ? "a" : undefined} />)}
          </BarChart>
        )}
      </ResponsiveContainer>
    </div>
  );
}

export function BlockView({ b }: { b: AnalysisBlock }) {
  if (b.type === "text") return <div className="text-[14px]"><Markdown>{b.markdown}</Markdown></div>;
  const title = b.title && <div className="mb-3 text-[14px] font-semibold">{b.title}</div>;
  if (b.type === "kpis")
    return (
      <div>
        {title}
        <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
          {b.items.map((it) => (
            <div key={it.label} className="rounded-2xl border bg-background/60 p-4">
              <div className="text-[12px] font-medium text-muted-foreground">{it.label}</div>
              <div className="mt-1 text-[26px] font-semibold tracking-tight tabular-nums">{formatValue(b.rows?.[0]?.[it.column], it.format)}</div>
            </div>
          ))}
        </div>
      </div>
    );
  if (!b.rows) return <div className="text-[13px] text-muted-foreground">No data (refresh the analysis).</div>;
  if (b.type === "table") return <div>{title}<DataTable rows={b.rows} columns={b.columns} /></div>;
  if (b.type === "chart") return <div>{title}<Chart b={b} /></div>;
  const p = pivotRows(b.rows, b.pivot);
  return <div>{title}<DataTable rows={p.lines as Rows} columns={[...b.pivot.rows, ...p.columns]} /></div>;
}
