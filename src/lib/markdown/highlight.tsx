/**
 * A small syntax highlighter for the code blocks of the markdown mockups (TypeScript, DAX, TMDL,
 * JSON, bash): comments, strings, keywords, numbers, calls and types, by regular expressions. No
 * library: the same tokens give the colors on screen and the language of Confluence's code macro.
 */
import type { ReactNode } from "react";

const KEYWORDS: Record<string, string[]> = {
  ts: "import export from const let var function return async await if else for of in new type interface class extends as default true false null undefined throw try catch".split(" "),
  dax: "EVALUATE DEFINE MEASURE VAR RETURN ORDER BY ASC DESC TOPN SUMMARIZECOLUMNS CALCULATE FILTER ALL VALUES SUM DIVIDE TREATAS KEEPFILTERS".split(" "),
  tmdl: "table column measure partition expression dataType formatString sourceColumn lineageTag mode source relationship fromColumn toColumn import".split(" "),
  json: ["true", "false", "null"],
  bash: "bun npm git cd export curl".split(" "),
};
const ALIASES: Record<string, string> = { tsx: "ts", typescript: "ts", js: "ts", javascript: "ts", sh: "bash", shell: "bash" };
export const langOf = (lang = "") => ALIASES[lang] ?? lang;

const COMMENT: Record<string, RegExp> = { ts: /\/\/[^\n]*|\/\*[\s\S]*?\*\//y, dax: /\/\/[^\n]*|--[^\n]*/y, tmdl: /\/\/[^\n]*/y, bash: /#[^\n]*/y, json: /$^/y };
const CLASSES = {
  comment: "text-zinc-400 italic",
  string: "text-emerald-700 dark:text-emerald-400",
  keyword: "text-violet-700 dark:text-violet-400 font-medium",
  number: "text-orange-700 dark:text-orange-400",
  call: "text-blue-700 dark:text-blue-400",
  type: "text-teal-700 dark:text-teal-400",
  key: "text-sky-800 dark:text-sky-300",
};

/** The code as colored spans (unknown languages: plain text). */
export function highlight(code: string, lang?: string): ReactNode[] {
  const l = langOf(lang);
  const words = new Set(KEYWORDS[l] ?? []);
  if (!KEYWORDS[l]) return [code];
  const out: ReactNode[] = [];
  let plain = "";
  const push = (text: string, kind: keyof typeof CLASSES) => {
    if (plain) out.push(plain);
    plain = "";
    out.push(
      <span key={out.length} className={CLASSES[kind]}>
        {text}
      </span>,
    );
  };
  const at = (re: RegExp, i: number) => ((re.lastIndex = i), re.exec(code)?.[0]);
  for (let i = 0; i < code.length; ) {
    const comment = at(COMMENT[l] ?? /$^/y, i);
    const string = at(/"(?:[^"\\\n]|\\.)*"|'(?:[^'\\\n]|\\.)*'|`(?:[^`\\]|\\.)*`/y, i);
    const word = at(/[A-Za-z_$][\w$]*/y, i);
    const number = at(/\d+(?:\.\d+)?(?:_\d+)*/y, i);
    if (comment) (push(comment, "comment"), (i += comment.length));
    else if (string) {
      // A JSON key: a string followed by ":".
      push(string, l === "json" && /^\s*:/.test(code.slice(i + string.length)) ? "key" : "string");
      i += string.length;
    } else if (word) {
      const prev = code[i - 1];
      const kind = words.has(word) || (l === "dax" && words.has(word.toUpperCase())) ? "keyword" : code[i + word.length] === "(" ? "call" : /^[A-Z]/.test(word) && l === "ts" && prev !== "." ? "type" : undefined;
      if (kind) push(word, kind);
      else plain += word;
      i += word.length;
    } else if (number && !/\w/.test(code[i - 1] ?? "")) (push(number, "number"), (i += number.length));
    else plain += code[i++];
  }
  if (plain) out.push(plain);
  return out;
}
