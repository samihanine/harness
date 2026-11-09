/** Tests run in Bun with a DOM (the Confluence converter parses HTML; the screens' code expects window). */
import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();
