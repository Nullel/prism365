import { type Policy, policySchema, requiresMfa } from "./conditional-access.ts";
import { policyReach } from "./policy-reach.ts";

// Whether an enabled Conditional Access policy that asks for MFA reaches one person. Who a policy
// reaches is decided once, by policyReach (the person page's Access tab uses the same function);
// this file only picks the policies that count and applies the "no other policy covers them" rule.
// It works over the person's ID, direct groups and roles: nested groups are not followed, and apps,
// locations, platforms and risk levels are not weighed, the same limits the Conditional Access
// audit states for its own MFA check.

export interface MfaPolicy {
  id: string;
  name: string;
  policy: Policy;
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
    if (policy.state.toLowerCase() !== "enabled" || !requiresMfa(policy)) continue;
    policies.push({ id: policy.id, name: policy.displayName ?? "an unnamed policy", policy });
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

export type MfaReach =
  | { state: "covered" }
  | { state: "left-out"; policy: string }
  | { state: "none" };

// Covered when some policy includes the person and none of the policies that include them excludes
// them from that one. Left out when a policy would include them but names them in its exclusions
// and no other policy covers them. Otherwise none: no MFA policy is aimed at them at all, which
// is a different finding (the Conditional Access audit's) and not a combination.
export function mfaReach(who: ReachSubject, policies: readonly MfaPolicy[]): MfaReach {
  const left: string[] = [];
  const person = {
    id: who.id,
    userType: who.guest ? "Guest" : "Member",
    groupIds: who.groupIds ?? [],
    roleTemplateIds: who.roleIds,
  };
  for (const { policy, name } of policies) {
    const reach = policyReach(policy, person);
    if (!reach.targeted) continue;
    if (reach.effect === "excluded") left.push(name);
    else return { state: "covered" };
  }
  const first = left.sort((a, b) => a.localeCompare(b))[0];
  return first === undefined ? { state: "none" } : { state: "left-out", policy: first };
}
