import { createFileRoute } from "@tanstack/react-router";
import { AgentPanel } from "@/components/agent-panel";

export const Route = createFileRoute("/")({ component: AgentPanel });
