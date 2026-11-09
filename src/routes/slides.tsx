import { createFileRoute } from "@tanstack/react-router";
import { ScreenHost } from "@/app/screen-host";
import { itemSearch } from "@/lib/features";
import { presenter } from "@/features/presenter/module";

export const Route = createFileRoute("/slides")({ validateSearch: itemSearch("deck"), component: () => <ScreenHost feature={presenter} /> });
