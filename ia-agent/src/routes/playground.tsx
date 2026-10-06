import { useEffect, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { PostMessageTransport } from "@modelcontextprotocol/ext-apps";
import { createDemoServer } from "@/playground/demo-server";
import type { Task } from "@/playground/demo-server";

export const Route = createFileRoute("/playground")({ component: Playground });

const initial: Task[] = [
  { id: "a1", title: "Write the quarterly report", status: "todo" },
  { id: "b2", title: "Review pull requests", status: "doing" },
  { id: "c3", title: "Book the team offsite", status: "todo" },
];

/** A demo host app: a task list on the left, the agent (iframe) on the right. */
function Playground() {
  const [tasks, setTasks] = useState(initial);
  const state = useRef(tasks);
  const iframe = useRef<HTMLIFrameElement>(null);

  useEffect(() => {
    const target = iframe.current!.contentWindow!;
    const server = createDemoServer({
      get: () => state.current,
      set: (next) => {
        state.current = next;
        setTasks(next);
      },
    });
    void server.connect(new PostMessageTransport(target, target));
    return () => void server.close();
  }, []);

  return (
    <div className="flex h-dvh">
      <section className="flex-1 overflow-y-auto p-8">
        <h1 className="mb-4 text-[15px] font-medium">Tasks</h1>
        <ul className="flex max-w-lg flex-col divide-y rounded-lg border bg-card shadow-soft">
          {tasks.map((task) => (
            <li key={task.id} className="flex items-center gap-3 px-3 py-2">
              <span className="font-mono text-[11px] text-muted-foreground">{task.id}</span>
              <span className="flex-1">{task.title}</span>
              <span className="rounded-full bg-muted px-2 py-0.5 text-[11px] text-muted-foreground">{task.status}</span>
            </li>
          ))}
        </ul>
      </section>
      <iframe ref={iframe} title="Agent" src={`${location.pathname}#/`} className="w-[420px] border-l" />
    </div>
  );
}
