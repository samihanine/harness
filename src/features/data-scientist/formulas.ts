/**
 * The calculation sheets' formulas, computed in the app as Excel computes them once downloaded: the
 * subset the AI is told — + - * / ^ & and comparisons, cells and ranges (A1, $B$2, A1:B4, 'Sheet'!C3),
 * SUM AVERAGE MIN MAX COUNT ROUND ABS IF IFERROR AND OR NOT, and GETPIVOTDATA("value", "pivot sheet",
 * "field", "item", …) to read a pivot's figure whatever its layout. Errors come back as Excel's
 * (#DIV/0!, #REF!, #NAME?, #VALUE!) with their reason.
 */
import { pivotData, type CalcSheet, type ChartSheet, type PivotSheet, type Sheet } from "./research";

export type Value = number | string | boolean | null;
/** A cell's result: its value, or an error ("#REF!") and why. */
export type Result = { value: Value; error?: string; why?: string };

class FormulaError extends Error {
  constructor(
    public code: string,
    why: string,
  ) {
    super(why);
  }
}

/** "B3" → [row 2, column 1]. */
export function cellIndex(ref: string): [number, number] {
  const m = ref.replace(/\$/g, "").toUpperCase().match(/^([A-Z]{1,3})(\d+)$/);
  if (!m) throw new FormulaError("#REF!", `"${ref}" is not a cell`);
  const column = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  return [Number(m[2]) - 1, column];
}
/** Column letters of a 0-based index (0 → A). */
export const columnName = (i: number): string => (i >= 26 ? columnName(Math.floor(i / 26) - 1) : "") + String.fromCharCode(65 + (i % 26));

type Token = { t: "num" | "str" | "ref" | "name" | "op" | "(" | ")" | "," | ":"; v: string };

function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) i++;
    else if (/[0-9.]/.test(c)) {
      const m = src.slice(i).match(/^\d*\.?\d+(?:[eE][-+]?\d+)?%?/)!;
      out.push({ t: "num", v: m[0] });
      i += m[0].length;
    } else if (c === '"') {
      let j = i + 1;
      let s = "";
      while (j < src.length && !(src[j] === '"' && src[j + 1] !== '"')) s += src[j] === '"' ? (j++, '"') : src[j], j++;
      out.push({ t: "str", v: s });
      i = j + 1;
    } else if (c === "'") {
      // 'Sheet name'!A1
      const j = src.indexOf("'!", i + 1);
      if (j < 0) throw new FormulaError("#REF!", "a sheet name opened with ' is not closed by '!");
      const m = src.slice(j + 2).match(/^\$?[A-Za-z]{1,3}\$?\d+/);
      if (!m) throw new FormulaError("#REF!", `no cell after ${src.slice(i, j + 2)}`);
      out.push({ t: "ref", v: `${src.slice(i + 1, j)}!${m[0]}` });
      i = j + 2 + m[0].length;
    } else if (/[A-Za-z_$]/.test(c)) {
      const m = src.slice(i).match(/^[A-Za-z_][A-Za-z0-9_.]*!\$?[A-Za-z]{1,3}\$?\d+|^\$?[A-Za-z]{1,3}\$?\d+(?![A-Za-z0-9_(])|^[A-Za-z_][A-Za-z0-9_.]*/)!;
      const word = m[0];
      out.push({ t: /^\$?[A-Za-z]{1,3}\$?\d+$/.test(word) || word.includes("!") ? "ref" : "name", v: word });
      i += word.length;
    } else if ("+-*/^&=".includes(c)) out.push({ t: "op", v: c }), i++;
    else if (c === "<" || c === ">") {
      const two = src.slice(i, i + 2);
      const op = two === "<=" || two === ">=" || two === "<>" ? two : c;
      out.push({ t: "op", v: op });
      i += op.length;
    } else if (c === "(" || c === ")" || c === "," || c === ";" || c === ":") out.push({ t: c === ";" ? "," : (c as Token["t"]), v: c }), i++;
    else throw new FormulaError("#NAME?", `unexpected "${c}"`);
  }
  return out;
}

/** What a formula reads: a cell of a sheet, a pivot's figure. */
type Reader = { cell: (sheet: string | undefined, row: number, column: number) => Value; pivot: (value: string, sheet: string, pairs: [string, Value][]) => number };

const num = (v: Value | Value[]): number => {
  if (Array.isArray(v)) throw new FormulaError("#VALUE!", "a range where one value is expected");
  if (v === null || v === "") return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  const n = typeof v === "number" ? v : Number(v);
  if (!Number.isFinite(n)) throw new FormulaError("#VALUE!", `"${v}" is not a number`);
  return n;
};
const flat = (args: (Value | Value[])[]) => args.flat().filter((v): v is number => typeof v === "number");

const FUNCTIONS: Record<string, (args: (Value | Value[])[], raw: () => (Value | Value[])[]) => Value> = {
  SUM: (a) => flat(a).reduce((x, y) => x + y, 0),
  AVERAGE: (a) => {
    const n = flat(a);
    if (!n.length) throw new FormulaError("#DIV/0!", "AVERAGE of no number");
    return n.reduce((x, y) => x + y, 0) / n.length;
  },
  MIN: (a) => (flat(a).length ? Math.min(...flat(a)) : 0),
  MAX: (a) => (flat(a).length ? Math.max(...flat(a)) : 0),
  COUNT: (a) => flat(a).length,
  ROUND: (a) => {
    const k = 10 ** num(a[1] ?? 0);
    return Math.round(num(a[0]) * k) / k;
  },
  ABS: (a) => Math.abs(num(a[0])),
  AND: (a) => a.flat().every((v) => !!num(v as Value)),
  OR: (a) => a.flat().some((v) => !!num(v as Value)),
  NOT: (a) => !num(a[0]),
};

/** A formula's parser and evaluator over a reader; IF and IFERROR evaluate their branches lazily. */
function evaluate(src: string, read: Reader, sheet: string): Value {
  const tokens = tokenize(src);
  let p = 0;
  const peek = () => tokens[p];
  const take = (t?: Token["t"], v?: string) => {
    const k = tokens[p];
    if (!k || (t && k.t !== t) || (v && k.v !== v)) throw new FormulaError("#NAME?", `expected ${v ?? t} near "${tokens.slice(p, p + 3).map((x) => x.v).join(" ")}"`);
    p++;
    return k;
  };
  const refValue = (ref: string): Value => {
    const [s, cell] = ref.includes("!") ? [ref.slice(0, ref.lastIndexOf("!")), ref.slice(ref.lastIndexOf("!") + 1)] : [undefined, ref];
    const [r, c] = cellIndex(cell);
    return read.cell(s ?? sheet, r, c);
  };
  const range = (from: string, to: string): Value[] => {
    const s = from.includes("!") ? from.slice(0, from.lastIndexOf("!")) : sheet;
    const [r1, c1] = cellIndex(from.slice(from.lastIndexOf("!") + 1));
    const [r2, c2] = cellIndex(to.slice(to.lastIndexOf("!") + 1));
    const out: Value[] = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.push(read.cell(s, r, c));
    return out;
  };
  // Expression grammar: comparison < concat < additive < multiplicative < power < unary < primary.
  const compare = (): Value | Value[] => {
    let left = concat();
    while (peek()?.t === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(peek().v)) {
      const op = take().v;
      const right = concat();
      const [a, b] = [left as Value, right as Value];
      const x = typeof a === "number" && typeof b === "number" ? [a, b] : [String(a ?? "").toLowerCase(), String(b ?? "").toLowerCase()];
      left = op === "=" ? x[0] === x[1] : op === "<>" ? x[0] !== x[1] : op === "<" ? x[0] < x[1] : op === ">" ? x[0] > x[1] : op === "<=" ? x[0] <= x[1] : x[0] >= x[1];
    }
    return left;
  };
  const concat = (): Value | Value[] => {
    let left = additive();
    while (peek()?.t === "op" && peek().v === "&") (take(), (left = `${left ?? ""}${additive() ?? ""}`));
    return left;
  };
  const additive = (): Value | Value[] => {
    let left = multiplicative();
    while (peek()?.t === "op" && (peek().v === "+" || peek().v === "-")) {
      const op = take().v;
      const right = multiplicative();
      left = op === "+" ? num(left) + num(right) : num(left) - num(right);
    }
    return left;
  };
  const multiplicative = (): Value | Value[] => {
    let left = power();
    while (peek()?.t === "op" && (peek().v === "*" || peek().v === "/")) {
      const op = take().v;
      const right = power();
      if (op === "/" && num(right) === 0) throw new FormulaError("#DIV/0!", "division by zero");
      left = op === "*" ? num(left) * num(right) : num(left) / num(right);
    }
    return left;
  };
  const power = (): Value | Value[] => {
    let left = unary();
    while (peek()?.t === "op" && peek().v === "^") (take(), (left = num(left) ** num(unary())));
    return left;
  };
  const unary = (): Value | Value[] => {
    if (peek()?.t === "op" && (peek().v === "-" || peek().v === "+")) return take().v === "-" ? -num(unary()) : num(unary());
    return primary();
  };
  /** The arguments of a call, each evaluated only when asked (IF, IFERROR). */
  const argsLazy = () => {
    take("(");
    const starts: number[] = [];
    let depth = 0;
    if (peek()?.t !== ")") starts.push(p);
    while (p < tokens.length) {
      const k = tokens[p];
      if (k.t === "(") depth++;
      else if (k.t === ")") {
        if (!depth) break;
        depth--;
      } else if (k.t === "," && !depth) starts.push(p + 1);
      p++;
    }
    take(")");
    const at = (i: number) => {
      if (i >= starts.length) return null;
      const save = p;
      p = starts[i];
      try {
        return compare();
      } finally {
        p = save;
      }
    };
    return { count: starts.length, at };
  };
  const primary = (): Value | Value[] => {
    const k = take();
    if (k.t === "num") return k.v.endsWith("%") ? Number(k.v.slice(0, -1)) / 100 : Number(k.v);
    if (k.t === "str") return k.v;
    if (k.t === "(") {
      const v = compare();
      take(")");
      return v;
    }
    if (k.t === "ref") {
      if (peek()?.t === ":") {
        take(":");
        return range(k.v, take("ref").v);
      }
      return refValue(k.v);
    }
    if (k.t === "name") {
      const name = k.v.toUpperCase();
      if (name === "TRUE" || name === "FALSE") return name === "TRUE";
      const a = argsLazy();
      if (name === "IF") return num(a.at(0) as Value) ? (a.at(1) ?? true) : (a.at(2) ?? false);
      if (name === "IFERROR") {
        try {
          return a.at(0);
        } catch (e) {
          if (e instanceof FormulaError) return a.at(1);
          throw e;
        }
      }
      const values = Array.from({ length: a.count }, (_, i) => a.at(i) as Value | Value[]);
      if (name === "GETPIVOTDATA") {
        const [value, target, ...rest] = values as Value[];
        const pairs: [string, Value][] = [];
        for (let i = 0; i + 1 < rest.length; i += 2) pairs.push([String(rest[i]), rest[i + 1]]);
        return read.pivot(String(value), String(target), pairs);
      }
      const f = FUNCTIONS[name];
      if (!f) throw new FormulaError("#NAME?", `unknown function ${name} (use ${[...Object.keys(FUNCTIONS), "IF", "IFERROR", "GETPIVOTDATA"].join(", ")})`);
      return f(values, () => values);
    }
    throw new FormulaError("#NAME?", `unexpected "${k.v}"`);
  };
  const v = compare();
  if (p < tokens.length) throw new FormulaError("#NAME?", `unexpected "${tokens[p].v}"`);
  if (Array.isArray(v)) throw new FormulaError("#VALUE!", "a range is not a value: SUM(A1:A3)…");
  return v;
}

/** A literal cell: a number when it reads as one ("12", "12.5", "-3", "8%"), else its text. */
const literal = (raw: string): Value => {
  const s = raw.trim();
  if (s === "") return null;
  if (/^-?\d+(\.\d+)?%$/.test(s)) return Number(s.slice(0, -1)) / 100;
  return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : raw;
};

/**
 * Every cell of the calculation sheets computed against the workbook's sheets: results by sheet name.
 * Pivot and chart sheets are read through GETPIVOTDATA; calculation sheets by their cells.
 */
export function computeWorkbook(sheets: Sheet[]) {
  const byName = new Map(sheets.map((s) => [s.name.toLowerCase(), s]));
  const results = new Map<string, Result[][]>();
  const busy = new Set<string>();
  const sheetOf = (name: string) => {
    const s = byName.get(name.toLowerCase());
    if (!s) throw new FormulaError("#REF!", `no sheet "${name}" (sheets: ${sheets.map((x) => x.name).join(", ")})`);
    return s;
  };
  const cellOf = (s: CalcSheet, r: number, c: number): Result => {
    const raw = s.cells[r]?.[c] ?? "";
    if (!raw.trim().startsWith("=")) return { value: literal(raw) };
    const key = `${s.name}!${r},${c}`;
    if (busy.has(key)) return { value: "#REF!", error: "#REF!", why: "circular reference" };
    busy.add(key);
    try {
      const value = evaluate(raw.trim().slice(1), reader, s.name);
      return { value: typeof value === "number" && !Number.isFinite(value) ? "#NUM!" : value };
    } catch (e) {
      const code = e instanceof FormulaError ? e.code : String(e instanceof Error ? e.message : e).match(/^#[A-Z/0!?]+/)?.[0] ?? "#VALUE!";
      return { value: code, error: code, why: e instanceof Error ? e.message.replace(/^#[A-Z/0!?]+\s*/, "") : String(e) };
    } finally {
      busy.delete(key);
    }
  };
  const reader: Reader = {
    cell: (name, r, c) => {
      const s = sheetOf(name ?? "");
      if (s.kind !== "calc") throw new FormulaError("#REF!", `"${s.name}" is a ${s.kind}: read it with GETPIVOTDATA`);
      const res = cellOf(s, r, c);
      if (res.error) throw new FormulaError(res.error, res.why ?? "");
      return res.value;
    },
    pivot: (value, name, pairs) => {
      const s = sheetOf(name);
      if (s.kind === "calc") throw new FormulaError("#REF!", `"${s.name}" is a calculation, not a pivot`);
      try {
        const n = pivotData(s as PivotSheet | ChartSheet, value, pairs);
        if (n === undefined) throw new FormulaError("#REF!", `"${s.name}" has no rows yet (refresh it)`);
        return n;
      } catch (e) {
        throw e instanceof FormulaError ? e : new FormulaError("#REF!", e instanceof Error ? e.message.replace(/^#REF!\s*/, "") : String(e));
      }
    },
  };
  for (const s of sheets) if (s.kind === "calc") results.set(s.name, s.cells.map((line, r) => line.map((_, c) => cellOf(s, r, c))));
  return results;
}
