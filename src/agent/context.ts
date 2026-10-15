/** Pieces of the AI context shared by the screens. */
import { contextsText } from "@/lib/library";

/**
 * The context of a screen with the user's context files linked to its entries (company vocabulary,
 * rules…), given once at the start of a conversation (again only when it changes).
 */
export async function withContexts(parts: (string | false | undefined)[], contextIds: (string[] | undefined)[]) {
  const ids = [...new Set(contextIds.flatMap((x) => x ?? []))];
  const contexts = ids.length ? await contextsText(ids) : "";
  return [...parts, contexts && `Context files from the user (background knowledge; follow their rules):\n${contexts}`].filter(Boolean).join("\n\n");
}
