import { createFileRoute } from "@tanstack/react-router";
import { ScreenHost } from "@/app/screen-host";
import { itemSearch } from "@/lib/features";
import { reports } from "@/features/reports/module";

export const Route = createFileRoute("/reports")({ validateSearch: itemSearch("report"), component: () => <ScreenHost feature={reports} /> });
