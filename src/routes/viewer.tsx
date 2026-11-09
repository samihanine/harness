import { createFileRoute } from "@tanstack/react-router";
import { ScreenHost } from "@/app/screen-host";
import { itemSearch } from "@/lib/features";
import { trainer } from "@/features/trainer/module";

export const Route = createFileRoute("/viewer")({ validateSearch: itemSearch("report"), component: () => <ScreenHost feature={trainer} /> });
