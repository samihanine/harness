/** The AI asks the user: what it found and its questions, each with the real options as buttons. Ends the turn. */
import { z } from "zod";
import { tool } from "../tool";

export const askUser = tool({
  name: "ask_user",
  description:
    "Ask the user to clarify or confirm before acting. Say briefly what you found (message), then 1 to 3 questions, each with 2 to 6 short options built from what you explored (real names and values, the most likely first). The user answers with a click or in words. Your turn ends with this call: write nothing after it.",
  args: z.object({
    message: z.string().optional().describe("What you found / understood, in 1-3 short sentences (markdown)."),
    questions: z
      .array(z.object({ question: z.string(), options: z.array(z.string()).min(2).max(6), multiple: z.boolean().optional().describe("Several options can be chosen.") }))
      .min(1)
      .max(3),
  }),
  async run(args, ctx) {
    ctx.conversation.scratch.askedPending = true;
    ctx.effect("ask", args);
    return "The questions are shown to the user. Stop now: the answer comes with the next message.";
  },
});
