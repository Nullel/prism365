import type { Severity } from "@prism365/contracts";
import type { Verdict } from "./account-activity.ts";
import type { InputId, Person } from "./combined-risks-inputs.ts";
import { type MfaPolicy, mfaReach } from "./combined-risks-reach.ts";
import { complianceLabel } from "./devices.ts";

// The rules. Each is a pure function over one person's joined inputs: it returns the reason, one
// sentence naming the facts, or null. A rule trips only on something an input says, never on a
// person being absent from an input. Which inputs a rule needs, and which of them can stand in for
// each other, is declared with it; the audit decides from that whether the rule can be checked.

export const QUIET_SIGN_IN_DAYS = 30;
export const OLD_PASSWORD_DAYS = 365;

export type RuleId =
  | "admin-unprotected"
  | "risky-unprotected"
  | "admin-signin-finding"
  | "dormant-admin"
  | "guest-with-power"
  | "weak-password-no-mfa"
  | "admin-weak-device"
  | "disabled-still-holding"
  | "paying-for-nothing";

export interface RuleEnv {
  now: Date;
  // The enabled Conditional Access policies that require MFA, or null when they were not read.
  policies: readonly MfaPolicy[] | null;
}

export interface Trip {
  reason: string;
  // The reason rests on the Conditional Access policies, read live for this run.
  usedPolicies?: boolean;
}

export interface Rule {
  id: RuleId;
  label: string;
  severity: Extract<Severity, "critical" | "high" | "medium">;
  // Every group must be met by at least one of its inputs, which are tried in order. A rule with
  // a group nothing can meet is not checked for anyone.
  needs: readonly (readonly InputId[])[];
  check(who: Person, env: RuleEnv): Trip | null;
}

const DAY = 86_400_000;

// Typed as Verdict and labelled by the devices audit, so a renamed value fails the build instead
// of quietly matching nothing.
const QUIET: Verdict = "Quiet everywhere";
const NEVER: Verdict = "Never signed in";
const NOT_COMPLIANT = complianceLabel("noncompliant");
const MEDIUM_RANK = 2;
const severityRank: Record<string, number> = { critical: 4, high: 3, medium: MEDIUM_RANK };

const privilegedHolds = (who: Person) => who.roles.filter((role) => role.privileged);

function roleText(roles: readonly { role: string; held: string }[]): string {
  const names = [...new Set(roles.map((item) => `${item.role} (${item.held.toLowerCase()})`))];
  if (names.length <= 2) return names.join(" and ");
  return `${names.slice(0, 2).join(", ")} and ${names.length - 2} more`;
}

const plural = (count: number, one: string, many: string) => `${count} ${count === 1 ? one : many}`;

// Role template IDs of the privileged roles a person holds, which is how policies name roles.
const roleIds = (who: Person) =>
  who.roles.flatMap((role) => (role.templateId ? [role.templateId] : []));

// What stands between a person and MFA: no registration, or a policy that requires it and leaves
// the person out. null when neither is shown.
function unprotected(who: Person, env: RuleEnv): { text: string; usedPolicies: boolean } | null {
  const parts: string[] = [];
  let usedPolicies = false;
  if (who.mfa === false) parts.push("has no MFA registered");
  if (env.policies) {
    const reach = mfaReach(
      { id: who.id, groupIds: who.groupIds, roleIds: roleIds(who), guest: who.guest },
      env.policies,
    );
    if (reach.state === "left-out") {
      parts.push(`is left out of ${reach.policy}, which requires MFA`);
      usedPolicies = true;
    }
  }
  return parts.length === 0 ? null : { text: parts.join(" and "), usedPolicies };
}

// Whole days between two instants, never negative.
const daysBetween = (from: string, to: Date) =>
  Math.max(0, Math.floor((to.getTime() - Date.parse(from)) / DAY));

export const RULES: readonly Rule[] = [
  {
    id: "admin-unprotected",
    label: "Administrator without MFA",
    severity: "critical",
    needs: [["roles"], ["mfa"]],
    check(who, env) {
      const held = privilegedHolds(who);
      if (held.length === 0) return null;
      const gap = unprotected(who, env);
      if (!gap) return null;
      return {
        reason: `Holds ${roleText(held)} and ${gap.text}.`,
        ...(gap.usedPolicies ? { usedPolicies: true } : {}),
      };
    },
  },
  {
    id: "risky-unprotected",
    label: "At-risk account without MFA",
    severity: "critical",
    needs: [["risky"], ["mfa"]],
    check(who, env) {
      const risk = who.risk;
      if (!risk) return null;
      const level = risk.level.toLowerCase();
      if (level !== "high" && level !== "medium") return null;
      // Remediated and dismissed risk is closed; "at risk" and "confirmed compromised" are not.
      const state = risk.state.toLowerCase();
      if (state === "remediated" || state === "dismissed") return null;
      const gap = unprotected(who, env);
      if (!gap) return null;
      return {
        reason: `Is marked ${level} risk (${state}) and ${gap.text}.`,
        ...(gap.usedPolicies ? { usedPolicies: true } : {}),
      };
    },
  },
  {
    id: "admin-signin-finding",
    label: "Administrator with a sign-in finding",
    severity: "critical",
    needs: [["roles"]],
    check(who) {
      const held = privilegedHolds(who);
      if (held.length === 0) return null;
      // Medium and above: the audit's own threshold for a finding worth joining to an account.
      const worst = who.findings
        .filter((finding) => (severityRank[finding.severity] ?? 0) >= MEDIUM_RANK)
        .sort((a, b) => (severityRank[b.severity] ?? 0) - (severityRank[a.severity] ?? 0))[0];
      if (!worst) return null;
      return {
        reason: `Has an open ${worst.severity} severity sign-in finding (${worst.title.slice(0, 120)}) and holds ${roleText(held)}.`,
      };
    },
  },
  {
    id: "dormant-admin",
    label: "Dormant administrator",
    severity: "high",
    needs: [["roles"], ["activity", "inactive"]],
    check(who, env) {
      const held = privilegedHolds(who);
      if (held.length === 0) return null;
      if (who.lastSuccess !== null) {
        const days = daysBetween(who.lastSuccess, env.now);
        return days >= QUIET_SIGN_IN_DAYS
          ? {
              reason: `Holds ${roleText(held)} and has had no successful sign-in for ${plural(days, "day", "days")}.`,
            }
          : null;
      }
      // Nothing recorded: the account's own age keeps a new one from counting, and an unreadable
      // age (null) proves nothing.
      if (who.days === null || who.days < QUIET_SIGN_IN_DAYS) return null;
      return {
        reason: `Holds ${roleText(held)} and has no successful sign-in on record, for ${plural(who.days, "day", "days")}.`,
      };
    },
  },
  {
    id: "guest-with-power",
    label: "Guest with directory power",
    severity: "high",
    needs: [["roles"], ["guests", "activity", "inactive", "mfa"]],
    check(who) {
      if (!who.guest) return null;
      const parts: string[] = [];
      if (who.roles.length > 0) parts.push(`holds ${roleText(who.roles)}`);
      if (who.ownedGroups.length > 0)
        parts.push(`owns ${plural(who.ownedGroups.length, "group", "groups")}`);
      return parts.length === 0
        ? null
        : { reason: `Is a guest account that ${parts.join(" and ")}.` };
    },
  },
  {
    id: "weak-password-no-mfa",
    label: "Weak password, no MFA",
    severity: "high",
    needs: [["passwords"], ["mfa"]],
    check(who) {
      const password = who.password;
      if (!password || who.mfa !== false) return null;
      const facts: string[] = [];
      if (password.neverExpires) facts.push("its password never expires");
      if (password.weak) facts.push("a strong password is not required");
      if (password.days !== null && password.days > OLD_PASSWORD_DAYS)
        facts.push(`its password is ${plural(password.days, "day", "days")} old`);
      if (facts.length === 0) return null;
      const lead = facts.join(", ").replace(/^./, (letter) => letter.toUpperCase());
      return { reason: `${lead}, and no MFA is registered.` };
    },
  },
  {
    id: "admin-weak-device",
    label: "Administrator on a weak device",
    severity: "high",
    needs: [["roles"], ["devices"]],
    check(who) {
      const held = privilegedHolds(who);
      if (held.length === 0) return null;
      const weak = who.devices
        .map((device) => ({
          name: device.name ?? "a managed device",
          problems: [
            ...(device.compliance === NOT_COMPLIANT ? ["not compliant"] : []),
            ...(device.encrypted === false ? ["not encrypted"] : []),
          ],
        }))
        .find((device) => device.problems.length > 0);
      if (!weak) return null;
      return {
        reason: `Holds ${roleText(held)} and uses ${weak.name}, which is ${weak.problems.join(" and ")}.`,
      };
    },
  },
  {
    id: "disabled-still-holding",
    label: "Disabled but still holding access",
    severity: "medium",
    needs: [["roles"], ["activity", "inactive", "reclaim"]],
    check(who) {
      if (who.enabled !== false) return null;
      const parts: string[] = [];
      if (who.roles.length > 0) parts.push(roleText(who.roles));
      if (who.licenses.length > 0) parts.push(plural(who.licenses.length, "license", "licenses"));
      if (who.ownedGroups.length > 0)
        parts.push(`ownership of ${plural(who.ownedGroups.length, "group", "groups")}`);
      return parts.length === 0
        ? null
        : { reason: `The account is disabled but still holds ${parts.join(", ")}.` };
    },
  },
  {
    id: "paying-for-nothing",
    label: "Paying for an unused account",
    severity: "medium",
    needs: [["activity"], ["mailbox"]],
    check(who) {
      if (who.licenses.length === 0) return null;
      if (who.verdict !== QUIET && who.verdict !== NEVER) return null;
      // A shared, room or equipment mailbox is quiet by design and may need its license.
      if (who.mailboxType && /shared|room|equipment/i.test(who.mailboxType)) return null;
      const how = who.verdict === NEVER ? "has never signed in" : "has been quiet everywhere";
      return {
        reason: `Holds ${plural(who.licenses.length, "license", "licenses")} and ${how}.`,
      };
    },
  },
];
