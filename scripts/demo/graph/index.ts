import { applications } from "./applications.ts";
import { conditionalAccess } from "./conditional-access.ts";
import { managedDevices } from "./devices.ts";
import { directoryAudits } from "./directory-audits.ts";
import { riskyUsers } from "./risky-users.ts";
import { directoryRoles } from "./roles.ts";
import type { DemoGraphExtension } from "./types.ts";
import { usageReports } from "./usage-reports.ts";

export type { DemoGraphExtension, DemoGraphRequest, DemoGraphTools } from "./types.ts";

// Every feature's part of the synthetic Graph, asked in order after the built-in paths.
export const demoGraphExtensions: readonly DemoGraphExtension[] = [
  directoryAudits,
  conditionalAccess,
  riskyUsers,
  managedDevices,
  applications,
  usageReports,
  directoryRoles,
];
