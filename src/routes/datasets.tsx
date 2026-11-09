import { createFileRoute } from "@tanstack/react-router";
import { ScreenHost } from "@/app/screen-host";
import { itemSearch } from "@/lib/features";
import { datasets } from "@/features/datasets/module";

export const Route = createFileRoute("/datasets")({ validateSearch: itemSearch("dataset"), component: () => <ScreenHost feature={datasets} /> });
