import type { PrismModule } from "@prism365/modules";
import {
  COLLECTOR_ID,
  type CollectorOptions,
  collectDirectoryChanges,
  MODULE_ID,
} from "./collector.ts";
import { type DirectoryChangesConfig, defaultDirectoryChangesConfig } from "./config.ts";
import { directoryChangesMigrations } from "./migrations.ts";

export * from "./collector.ts";
export * from "./config.ts";
export * from "./cursor.ts";
export * from "./migrations.ts";
export * from "./query.ts";
export * from "./rules.ts";
export * from "./source.ts";

// Directory changes: every Microsoft Entra directory audit event, kept for as long as Prism365
// runs, and findings for the changes that matter for security. See
// docs/modules/directory-changes.md.
export function createDirectoryChangesModule(
  config: DirectoryChangesConfig = defaultDirectoryChangesConfig,
  options: CollectorOptions = {},
): PrismModule {
  return {
    id: MODULE_ID,
    title: "Directory changes",
    description:
      "Keeps every Microsoft Entra directory audit event, who changed which account, group, role, app or policy and when, for as long as Prism365 runs, and flags the changes that matter for security.",
    version: 1,
    packs: ["audit.directory.read"],
    requires: ["role.reportsReader"],
    migrations: directoryChangesMigrations,
    collectors: [
      {
        id: COLLECTOR_ID,
        title: "Directory audit log",
        cadenceSeconds: 600,
        timeoutSeconds: 120,
        run: (context) => collectDirectoryChanges(context, config, options),
      },
    ],
    // No retention rule: Microsoft keeps these records 7 or 30 days, and keeping the history
    // after that is what the module is for. The operator deletes it (docs/PRIVACY.md).
    retention: [],
  };
}
