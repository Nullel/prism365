import { type AuditDefinition, auditDefinitionSchema } from "@prism365/contracts";
import { type GraphFeatureId, permissionPacks } from "@prism365/graph";
import { licenseCapacityAudit } from "./capacity.ts";
import type { Audit } from "./definition.ts";
import { guestAccountsAudit } from "./guests.ts";
import { inactiveAccountsAudit } from "./inactive.ts";
import { mfaRegistrationAudit } from "./mfa.ts";
import { licenseReclaimAudit } from "./reclaim.ts";

export * from "./definition.ts";
export * from "./directory.ts";
export * from "./scope.ts";
export * from "./store.ts";
export { strongestMethod, methodLabel } from "./mfa.ts";
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

export function findAudit(id: string): Audit | undefined {
  return builtInAudits.find((audit) => audit.definition.id === id);
}
