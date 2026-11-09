import { createFileRoute } from "@tanstack/react-router";
import { ScreenHost } from "@/app/screen-host";
import { itemSearch } from "@/lib/features";
import { builder } from "@/features/builder/module";

export const Route = createFileRoute("/editor")({ validateSearch: itemSearch("report"), component: () => <ScreenHost feature={builder} /> });
