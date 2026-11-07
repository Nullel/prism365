import type { Policy } from "./conditional-access.ts";

// Whether a Conditional Access policy reaches one person, from the "Users" part of the policy
// alone. Conditions such as apps, platforms, locations and risk decide whether it applies at a
// given sign-in, and are not evaluated here. Microsoft evaluates an exclusion after the inclusions
// and lets it win, so a person both included and excluded is left out.

export interface ReachPerson {
  id: string;
  // Microsoft's userType: "Member" or "Guest". null when it could not be read.
  userType: string | null;
  // Every group the person is in, directly or through another group.
  groupIds: readonly string[];
  // The role templates of the directory roles the person holds now.
  roleTemplateIds: readonly string[];
}

export interface ReachReason {
  kind: "all" | "person" | "group" | "role" | "guests";
  // The group, role template or account id; null for "all" and "guests".
  id: string | null;
}

export interface Reach {
  // "excluded" when any exclusion names the person, whether or not an inclusion reached them.
  effect: "applies" | "excluded" | "none";
  targeted: boolean;
  included: ReachReason[];
  excluded: ReachReason[];
}

const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

function matches(
  person: ReachPerson,
  rule: {
    users: readonly string[];
    groups: readonly string[];
    roles: readonly string[];
    guests: { types: readonly string[] } | null;
  },
  allowAll: boolean,
): ReachReason[] {
  const reasons: ReachReason[] = [];
  const guest = person.userType !== null && same(person.userType, "Guest");
  if (allowAll && rule.users.some((item) => same(item, "All")))
    reasons.push({ kind: "all", id: null });
  if (rule.users.some((item) => same(item, person.id)))
    reasons.push({ kind: "person", id: person.id });
  if (guest && rule.users.some((item) => same(item, "GuestsOrExternalUsers")))
    reasons.push({ kind: "guests", id: null });
  else if (guest && rule.guests !== null && rule.guests.types.length > 0)
    reasons.push({ kind: "guests", id: null });
  for (const id of rule.groups)
    if (person.groupIds.some((held) => same(held, id))) reasons.push({ kind: "group", id });
  for (const id of rule.roles)
    if (person.roleTemplateIds.some((held) => same(held, id))) reasons.push({ kind: "role", id });
  return reasons;
}

export function policyReach(policy: Policy, person: ReachPerson): Reach {
  const users = policy.conditions.users;
  const included = matches(
    person,
    {
      users: users.includeUsers,
      groups: users.includeGroups,
      roles: users.includeRoles,
      guests: users.includeGuestsOrExternalUsers,
    },
    true,
  );
  // "All" in an exclusion list is not a value Microsoft accepts, so it is not read as one.
  const excluded = matches(
    person,
    {
      users: users.excludeUsers,
      groups: users.excludeGroups,
      roles: users.excludeRoles,
      guests: users.excludeGuestsOrExternalUsers,
    },
    false,
  );
  const targeted = included.length > 0;
  return {
    effect: excluded.length > 0 ? "excluded" : targeted ? "applies" : "none",
    targeted,
    included,
    excluded,
  };
}
