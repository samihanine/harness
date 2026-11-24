import { saveFile } from "@server/files";
const csv = "Région;Responsable;Prime (%)\nEast;Alice Martin;2,5\nCentral;Bruno Leroy;3\nWest;Chloé Dubois;3,5\n";
console.log(JSON.stringify(await saveFile(new TextEncoder().encode(csv), "responsables.csv", "text/csv")));
