import type { PrismModule } from "@prism365/modules";
import { collectHealth, collectMessages, HEALTH_MODULE, MESSAGES_MODULE } from "./collector.ts";
import { messageCenterMigrations, serviceHealthMigrations } from "./migrations.ts";

export * from "./collector.ts";
export * from "./migrations.ts";
export * from "./source.ts";
export * from "./status.ts";

// Microsoft 365 service health: the services the tenant has, and Microsoft's incidents and
// advisories for them. Needs the service.health.read pack (ServiceHealth.Read.All).
export function createServiceHealthModule(): PrismModule {
  return {
    id: HEALTH_MODULE,
    title: "Service health",
    description:
      "Reads Microsoft's incidents and advisories for the services this tenant has, and notifies on new ones and meaningful changes.",
    version: 1,
    packs: ["service.health.read"],
    migrations: serviceHealthMigrations,
    collectors: [
      {
        id: "health",
        title: "Service health",
        cadenceSeconds: 300,
        timeoutSeconds: 120,
        run: (context) => collectHealth(context),
      },
    ],
    retention: [{ table: "issue", timeColumn: "last_modified_at", days: 180 }],
  };
}

// Message center: Microsoft's planned changes, retirements and required actions. A separate
// module so it can be enabled on its own; needs service.messages.read (ServiceMessage.Read.All).
export function createMessageCenterModule(): PrismModule {
  return {
    id: MESSAGES_MODULE,
    title: "Message center",
    description:
      "Reads Message center posts and keeps the ones with deadlines, retirements and major changes in Admin attention.",
    version: 1,
    packs: ["service.messages.read"],
    migrations: messageCenterMigrations,
    collectors: [
      {
        id: "messages",
        title: "Message center",
        cadenceSeconds: 1_800,
        timeoutSeconds: 120,
        run: (context) => collectMessages(context),
      },
    ],
    retention: [{ table: "post", timeColumn: "last_modified_at", days: 365 }],
  };
}
