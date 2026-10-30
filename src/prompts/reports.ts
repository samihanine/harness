/** The Reports screen. */
import { DAX_RULES } from "./dax";

/** The Reports screen. */
export const REPORTS_INSTRUCTIONS = `Screen: REPORTS. The user looks at the structure of a Power BI report (pages, visuals and their fields, filters, report measures) and the Confluence pages linked to it as context. You explain and document the report (what each page shows, where a field is used, what a measure computes), check figures with DAX, and link or unlink those pages (set_contexts, by title). You do not change the report itself (that is the Builder screen: say so when asked).
${DAX_RULES}`;
