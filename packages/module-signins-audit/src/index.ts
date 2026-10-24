import type { PrismModule } from "@prism365/modules";
import { CANARY_SETTING } from "./canaries.ts";
import { collectSignins, MODULE_ID } from "./collector.ts";
import { defaultSigninsAuditConfig, type SigninsAuditConfig } from "./config.ts";
import { signinsAuditMigrations } from "./migrations.ts";

export * from "./canaries.ts";
export * from "./collector.ts";
export * from "./config.ts";
export * from "./cursor.ts";
export * from "./migrations.ts";
export * from "./people.ts";
export * from "./rules.ts";
export * from "./source.ts";
export * from "./summary.ts";

// Sign-in audit: where accounts sign in from, what is normal for each, and which sign-ins and
// failure patterns break that. Security evidence only; see docs/modules/signins-audit.md.
export function createSigninsAuditModule(
  config: SigninsAuditConfig = defaultSigninsAuditConfig,
): PrismModule {
  return {
    id: MODULE_ID,
    title: "Sign-in audit",
    description:
      "Collects interactive sign-ins, learns each account's normal locations and devices, and flags sign-ins and failure patterns that break them.",
    version: 1,
    packs: ["audit.signins.read"],
    requires: ["entra.p1", "role.reportsReader"],
    // The canary list an administrator keeps through the API.
    settings: [CANARY_SETTING],
    migrations: signinsAuditMigrations,
    collectors: [
      {
        id: "signins",
        title: "Sign-in logs",
        cadenceSeconds: 300,
        timeoutSeconds: 240,
        run: (context) => collectSignins(context, config),
      },
    ],
    retention: [
      { table: "event", timeColumn: "created_at", days: 90 },
      { table: "seen", timeColumn: "last_seen", days: 365 },
    ],
  };
}
