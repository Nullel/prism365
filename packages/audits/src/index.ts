import { type AuditDefinition, auditDefinitionSchema } from "@prism365/contracts";
import { type GraphFeatureId, permissionPacks } from "@prism365/graph";
import { appConsentsAudit } from "./app-consents.ts";
import { licenseCapacityAudit } from "./capacity.ts";
import { conditionalAccessAudit } from "./conditional-access.ts";
import type { Audit } from "./definition.ts";
import { managedDevicesAudit } from "./devices.ts";
import { guestAccountsAudit } from "./guests.ts";
import { inactiveAccountsAudit } from "./inactive.ts";
import { m365AppUsageAudit } from "./m365-app-usage.ts";
import { mfaRegistrationAudit } from "./mfa.ts";
import { licenseReclaimAudit } from "./reclaim.ts";
import { riskyUsersAudit } from "./risky-users.ts";

// The devices, app consents and Microsoft 365 app usage audits, and what they share.
export {
  appConsentsAudit,
  HIGH_RISK_PERMISSIONS,
  highRiskIn,
  MAX_REQUESTS,
  MAX_SERVICE_PRINCIPALS,
  MICROSOFT_TENANTS,
  ownerOf,
} from "./app-consents.ts";
export { conditionalAccessAudit, createConditionalAccessAudit } from "./conditional-access.ts";
export * from "./definition.ts";
export {
  checkInOf,
  complianceLabel,
  jailbrokenLabel,
  MAX_DEVICES,
  managedDevicesAudit,
  ownershipLabel,
} from "./devices.ts";
export * from "./directory.ts";
export { firstAccepted, readPages } from "./graph-pages.ts";
export {
  createM365AppUsageAudit,
  lastActivityOf,
  MAX_USAGE_ACCOUNTS,
  m365AppUsageAudit,
  PERIOD_DAYS,
  usageOf,
} from "./m365-app-usage.ts";
export { methodLabel, strongestMethod } from "./mfa.ts";
export * from "./paced-read.ts";
export { createRiskyUsersAudit, riskyUsersAudit } from "./risky-users.ts";
export * from "./schedule.ts";
export * from "./scope.ts";
export * from "./store.ts";
export * from "./trend.ts";
export {
  guestAccountsAudit,
  inactiveAccountsAudit,
  licenseCapacityAudit,
  licenseReclaimAudit,
  mfaRegistrationAudit,
};

// Built-in audits, in the order the page lists them. A new audit is a new entry here.
export const builtInAudits: readonly Audit[] = [
  inactiveAccountsAudit,
  licenseReclaimAudit,
  mfaRegistrationAudit,
  guestAccountsAudit,
  licenseCapacityAudit,
  conditionalAccessAudit,
  riskyUsersAudit,
  managedDevicesAudit,
  appConsentsAudit,
  m365AppUsageAudit,
];

function permissionsOf(packs: readonly GraphFeatureId[]): string[] {
  return [
    ...new Set(
      packs.map((id) => {
        const pack = permissionPacks.find((candidate) => candidate.id === id);
        if (!pack) throw new Error(`Audit uses unknown read pack ${id}.`);
        return pack.permission;
      }),
    ),
  ];
}

// The published definition: the audit's own description plus the permissions its packs need.
export function auditDefinition(audit: Audit): AuditDefinition {
  const permissions = permissionsOf(audit.packs);
  return auditDefinitionSchema.parse({
    ...audit.definition,
    scoped: audit.scoped,
    permissions,
    optionalPermissions: permissionsOf(audit.optionalPacks).filter(
      (permission) => !permissions.includes(permission),
    ),
  });
}

// What one run may read: the audit's required packs plus only the optional packs the operator
// enabled, and only the granted permissions those packs carry. Several packs share a permission
// (AuditLog.Read.All serves sign-ins, directory audits and MFA registration), so a permission
// granted for one enabled pack must not open another pack the operator left off.
export function auditReadScope(
  audit: Audit,
  enabledPacks: readonly string[],
  grantedScopes: readonly string[],
): { packs: GraphFeatureId[]; granted: string[] } {
  const enabled = new Set(enabledPacks);
  const packs = [
    ...new Set([...audit.packs, ...audit.optionalPacks.filter((pack) => enabled.has(pack))]),
  ];
  const carried = new Set(permissionsOf(packs.filter((pack) => enabled.has(pack))));
  return { packs, granted: grantedScopes.filter((scope) => carried.has(scope)) };
}

export function findAudit(id: string): Audit | undefined {
  return builtInAudits.find((audit) => audit.definition.id === id);
}
