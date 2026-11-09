import { createFileRoute } from "@tanstack/react-router";
import { ScreenHost } from "@/app/screen-host";
import { itemSearch } from "@/lib/features";
import { excel } from "@/features/excel/module";

export const Route = createFileRoute("/excel")({ validateSearch: itemSearch("table"), component: () => <ScreenHost feature={excel} /> });
