/** DAX: the rules, the query tool, the performance limits the AI is told. */
import { PERFORMANCE } from "@/config";

/** Rules of every DAX query the AI writes. */
export const DAX_RULES = `DAX rules (follow them, queries fail otherwise):
- Use ONLY tables, columns and measures that exist, with their exact names copied from the model (context, describe_table). Never guess, shorten or translate a name. Not sure a column exists or in which table it is? search_model first.
- Syntax: 'Table'[Column], [Measure], text in "double quotes". Count rows with COUNTROWS('Table') or an existing measure.
- Grouping: SUMMARIZECOLUMNS('T'[a], 'T'[b], "Label", [Measure]). A plain list of value pairs of ONE table: DISTINCT(SELECTCOLUMNS('T', "A", 'T'[a], "B", 'T'[b])). Columns of two tables together only if they are related (see the relationships).
- ORDER BY only columns of the result. Count the parentheses before sending.
- Plain figures (values by groups, filtered, top N): the query tool writes the DAX for you. Raw DAX (run_dax) for the rest: time intelligence, filters on a measure, calculations.
- Limits of the app (to keep the models fast for everyone): a query returns at most a fixed number of rows (more are cut, the result says so) and is stopped after a few seconds. Aggregate (SUMMARIZECOLUMNS, TOPN) rather than reading raw rows; to know how many rows a query would return, EVALUATE ROW("n", COUNTROWS(<its table expression>)) first.
- Names written by the user (a pasted table, a screenshot, a column header of a report) are labels, not model names: find the model's real measure or column (context, search_model) before using it.
- A failed query: read the error, it names the problem (a missing name, or a (line, column) position) — fix exactly that and retry. To explain a failure to the user, quote the error; never invent a cause.`;

/** The limits of a query, as the AI reads them (read at call time: Settings › Performance changes them). */
export const daxLimits = () => `Limits: at most ${PERFORMANCE.maxRows} rows per query (more are cut), stopped after ${PERFORMANCE.timeoutMs / 1000} s.`;

/** What the `query` tool takes. */
export const QUERY_HELP = `query: a data question as a spec; the app writes and runs the DAX (names checked, the DAX used is returned). Prefer it for plain figures: values by groups, filtered, top N.
Its limits — use run_dax (raw DAX) for anything else: time intelligence (vs last year, YTD, moving averages), a filter on a measure's value, ratios or calculations between values, several tables not related, DAX functions of your own.`;
