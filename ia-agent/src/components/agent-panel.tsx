import { useEffect, useRef, useState } from "react";
import { DownloadIcon, PlugIcon, Settings2Icon, SquarePenIcon, UnplugIcon } from "lucide-react";
import { IconButton } from "@/components/icon-button";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import { downloadJson } from "@/lib/download";
import { conversations } from "@/lib/store";
import type { AgentItem } from "@/lib/types";
import { useAgent } from "@/lib/use-agent";
import { Composer } from "./composer";
import { HistoryMenu } from "./history-menu";
import { Message } from "./message";
import { SettingsDialog } from "./settings-dialog";
import { TraceDialog } from "./trace-dialog";

export function AgentPanel() {
  const agent = useAgent();
  const { host, conversation } = agent;
  const [draft, setDraft] = useState("");
  const [trace, setTrace] = useState<AgentItem | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const bottom = useRef<HTMLDivElement>(null);

  // Follow the conversation while it grows.
  const last = conversation.items.at(-1);
  useEffect(() => {
    void bottom.current?.scrollIntoView({ block: "end" });
  }, [conversation.items.length, last]);

  // Keep the trace dialog live while its run is going.
  const traced = trace && (conversation.items.find((i) => i.id === trace.id) as AgentItem | undefined);

  const exportAll = async () => {
    const all = await Promise.all(agent.list.map((c) => conversations.get(c.id)));
    downloadJson(all.filter(Boolean), `conversations-${new Date().toISOString().slice(0, 10)}`);
  };

  return (
    <div className="flex h-dvh flex-col bg-background">
      <header className="flex h-10 shrink-0 items-center gap-1 border-b px-1.5">
        <HistoryMenu
          title={conversation.title}
          list={agent.list}
          currentId={conversation.id}
          onOpen={(id) => void agent.open(id)}
          onRemove={(id) => void agent.remove(id)}
        />
        <span className="ml-auto" />
        <Connection host={host} />
        <IconButton label="New conversation" onClick={agent.newConversation} disabled={agent.running}>
          <SquarePenIcon />
        </IconButton>
        <DropdownMenu>
          <DropdownMenuTrigger render={<IconButton label="Export" />}>
            <DownloadIcon />
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="w-52">
            <DropdownMenuItem onClick={() => downloadJson(conversation, conversation.title)}>This conversation</DropdownMenuItem>
            <DropdownMenuItem onClick={() => void exportAll()}>All conversations of this app</DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
        <IconButton label="Settings" onClick={() => setSettingsOpen(true)}>
          <Settings2Icon />
        </IconButton>
      </header>

      <main className="min-h-0 flex-1 overflow-y-auto">
        {conversation.items.length === 0 ? (
          <Empty host={host} onPick={setDraft} />
        ) : (
          <div className="mx-auto flex max-w-3xl flex-col gap-5 px-4 py-5">
            {conversation.items.map((item) => (
              <Message
                key={item.id}
                item={item}
                canReply={!agent.running && item === last}
                onDecide={agent.decide}
                onTrace={setTrace}
                onUndo={(i) => void agent.undo(i)}
                onReply={(text) => void agent.send(text)}
              />
            ))}
            <div ref={bottom} />
          </div>
        )}
      </main>

      <footer className="mx-auto w-full max-w-3xl shrink-0 px-3 pb-3">
        <Composer
          running={agent.running}
          model={agent.model}
          onModel={agent.setModel}
          onSend={(text) => void agent.send(text)}
          onStop={agent.stop}
          draft={draft}
          onDraft={setDraft}
        />
      </footer>

      <TraceDialog item={traced ?? null} onClose={() => setTrace(null)} />
      <SettingsDialog open={settingsOpen} onClose={() => setSettingsOpen(false)} />
    </div>
  );
}

function Connection({ host }: { host: ReturnType<typeof useAgent>["host"] }) {
  const connected = host.status === "connected";
  return (
    <Tooltip>
      <TooltipTrigger className="flex size-7 items-center justify-center text-muted-foreground">
        {connected ? <PlugIcon className="size-3.5" /> : <UnplugIcon className="size-3.5 opacity-50" />}
      </TooltipTrigger>
      <TooltipContent>
        {connected
          ? `${host.title}: ${host.tools} tools, ${host.resources} resources`
          : host.status === "connecting"
            ? "Connecting…"
            : "No app connected"}
      </TooltipContent>
    </Tooltip>
  );
}

function Empty({ host, onPick }: { host: ReturnType<typeof useAgent>["host"]; onPick: (text: string) => void }) {
  const prompts = host.status === "connected" ? host.prompts : [];
  return (
    <div className="mx-auto flex h-full max-w-md flex-col justify-end gap-3 px-4 pb-4 animate-in fade-in-0">
      <p className="text-[15px] font-medium tracking-tight">
        {host.status === "connected" ? `How can I help with ${host.title}?` : "How can I help?"}
      </p>
      {host.status === "standalone" && (
        <p className="text-[12px] text-muted-foreground">No app connected: the agent only has its built-in tools.</p>
      )}
      {prompts.length > 0 && (
        <div className="flex flex-col">
          {prompts.map((p) => (
            <button
              key={p.name}
              type="button"
              onClick={() => host.status === "connected" && void host.host.prompt(p.name).then(onPick)}
              className="rounded-md px-2 py-1.5 text-left text-[12px] text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
            >
              {p.title}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
