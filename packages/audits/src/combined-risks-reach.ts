import { cleanText } from "@prism365/contracts";
import { z } from "zod";

// Whether an enabled Conditional Access policy that asks for MFA reaches one person. A small,
// pure check over the person's ID, direct groups and roles: nested groups are not followed (the
// person page checks those live), and apps, locations, platforms and risk levels are not weighed,
// the same limits the Conditional Access audit states for its own MFA check.

const list = z
  .array(z.unknown())
  .nullish()
  .catch(null)
  .transform((value) =>
    (value ?? [])
      .filter((item): item is string => typeof item === "string")
      .slice(0, 5_000)
      .map((item) => item.slice(0, 100).toLowerCase()),
  );

const selection = z
  .object({})
  .passthrough()
  .nullish()
  .catch(null)
  .transform((value) => value !== null && value !== undefined);

const policySchema = z.object({
  id: z.string().min(1).max(200),
  displayName: z
    .string()
    .nullish()
    .catch(null)
    .transform((value) => (value ? cleanText(value.slice(0, 200)) : null)),
  state: z.string().max(100),
  conditions: z
    .object({
      users: z
        .object({
          includeUsers: list,
          excludeUsers: list,
          includeGroups: list,
          excludeGroups: list,
          includeRoles: list,
          excludeRoles: list,
          includeGuestsOrExternalUsers: selection,
          excludeGuestsOrExternalUsers: selection,
        })
        .nullish()
        .catch(null),
    })
    .nullish()
    .catch(null),
  grantControls: z
    .object({
      builtInControls: list,
      authenticationStrength: selection,
    })
    .nullish()
    .catch(null),
});

export interface MfaPolicy {
  id: string;
  name: string;
  includeUsers: string[];
  excludeUsers: string[];
  includeGroups: string[];
  excludeGroups: string[];
  includeRoles: string[];
  excludeRoles: string[];
  includesGuests: boolean;
  excludesGuests: boolean;
}

// The enabled policies that require MFA, from the raw policy list. Policies in report-only mode
// are evaluated but not enforced, so they protect nobody and are left out. A policy that cannot be
// read is counted, not guessed at.
export function mfaPolicies(items: readonly unknown[]): { policies: MfaPolicy[]; skipped: number } {
  const policies: MfaPolicy[] = [];
  let skipped = 0;
  for (const raw of items) {
    const parsed = policySchema.safeParse(raw);
    if (!parsed.success) {
      skipped += 1;
      continue;
    }
    const policy = parsed.data;
    if (policy.state.toLowerCase() !== "enabled") continue;
    const grant = policy.grantControls;
    const requiresMfa =
      grant !== null &&
      grant !== undefined &&
      (grant.builtInControls.includes("mfa") || grant.authenticationStrength);
    if (!requiresMfa) continue;
    const users = policy.conditions?.users;
    if (!users) continue;
    policies.push({
      id: policy.id,
      name: policy.displayName ?? "an unnamed policy",
      includeUsers: users.includeUsers,
      excludeUsers: users.excludeUsers,
      includeGroups: users.includeGroups,
      excludeGroups: users.excludeGroups,
      includeRoles: users.includeRoles,
      excludeRoles: users.excludeRoles,
      includesGuests: users.includeGuestsOrExternalUsers,
      excludesGuests: users.excludeGuestsOrExternalUsers,
    });
  }
  return { policies, skipped };
}

export interface ReachSubject {
  id: string;
  // Direct groups; null when memberships were not read, and then no group is matched either way.
  groupIds: readonly string[] | null;
  // Role template IDs of the roles the person holds.
  roleIds: readonly string[];
  guest: boolean;
}

export type Reach =
  | { state: "covered" }
  | { state: "left-out"; policy: string }
  | { state: "none" };

const meets = (wanted: readonly string[], have: readonly string[] | null) => {
  if (!have) return false;
  const lower = new Set(have.map((item) => item.toLowerCase()));
  return wanted.some((item) => lower.has(item));
};

function includes(policy: MfaPolicy, who: ReachSubject): boolean {
  const users = policy.includeUsers;
  return (
    users.includes("all") ||
    users.includes(who.id.toLowerCase()) ||
    (who.guest && (users.includes("guestsorexternalusers") || policy.includesGuests)) ||
    meets(policy.includeGroups, who.groupIds) ||
    meets(policy.includeRoles, who.roleIds)
  );
}

function excludes(policy: MfaPolicy, who: ReachSubject): boolean {
  return (
    policy.excludeUsers.includes(who.id.toLowerCase()) ||
    (who.guest && policy.excludesGuests) ||
    meets(policy.excludeGroups, who.groupIds) ||
    meets(policy.excludeRoles, who.roleIds)
  );
}

// Covered when some policy includes the person and none of the policies that include them excludes
// them from that one. Left out when a policy would include them but names them in its exclusions
// and no other policy covers them. Otherwise none: no MFA policy is aimed at them at all, which
// is a different finding (the Conditional Access audit's) and not a combination.
export function mfaReach(who: ReachSubject, policies: readonly MfaPolicy[]): Reach {
  const left: string[] = [];
  for (const policy of policies) {
    if (!includes(policy, who)) continue;
    if (excludes(policy, who)) left.push(policy.name);
    else return { state: "covered" };
  }
  const first = left.sort((a, b) => a.localeCompare(b))[0];
  return first === undefined ? { state: "none" } : { state: "left-out", policy: first };
}
