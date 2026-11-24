import { createBlankReport } from "@server/pbi/new-report";
console.log(JSON.stringify(await createBlankReport(process.argv[2] ?? "Hub check", { id: "eda77871-2401-4255-ba6c-2fc3c66902a1" })));
