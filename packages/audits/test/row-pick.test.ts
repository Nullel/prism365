import { PGliteModuleDatabase } from "@prism365/storage";
import { rowPickSuite } from "./row-pick-suite.ts";

rowPickSuite("PGlite", () => PGliteModuleDatabase.open());
