import { createFileRoute } from "@tanstack/react-router";
import { ScreenHost } from "@/app/screen-host";
import { itemSearch } from "@/lib/features";
import { confluence } from "@/features/confluence/module";

export const Route = createFileRoute("/confluence")({ validateSearch: itemSearch("page"), component: () => <ScreenHost feature={confluence} /> });
