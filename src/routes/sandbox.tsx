import { createFileRoute } from "@tanstack/react-router";
import { ChatProposals } from "@/components/sandbox/chat-proposals";

/**
 * Sandbox (not in the menu): interface proposals on fake data, front end only. Now: the AI chat,
 * three proposals side by side, each as wide as the side panel.
 */
export const Route = createFileRoute("/sandbox")({ component: Sandbox });

function Sandbox() {
  return (
    <div className="flex h-full flex-col">
      <div className="flex h-11 shrink-0 items-center border-b bg-card px-4">
        <span className="text-[12px] font-semibold text-muted-foreground uppercase">AI chat · proposals</span>
        <span className="ml-3 text-[12px] text-muted-foreground">Same conversation, three designs; each frame is as wide as the side panel.</span>
      </div>
      <ChatProposals />
    </div>
  );
}
