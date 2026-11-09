import { createFileRoute } from "@tanstack/react-router";
import { ScreenHost } from "@/app/screen-host";
import { itemSearch } from "@/lib/features";
import { dataScientist } from "@/features/data-scientist/module";

export const Route = createFileRoute("/research")({ validateSearch: itemSearch("research"), component: () => <ScreenHost feature={dataScientist} /> });
