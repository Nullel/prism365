import type { Severity } from "@prism365/contracts";
import type { FindingInput } from "@prism365/modules";
import type { ChangeTarget, DirectoryChangeEvent } from "./source.ts";

// Rules over single directory changes. Each is pure: one stored event in, at most one finding
// out, keyed by the event id, so a change is flagged once however often it is read again.
//
// Activity and category names come from Microsoft's "Microsoft Entra audit log activity
// reference" (https://learn.microsoft.com/en-us/entra/identity/monitoring-health/reference-audit-activities,
// checked 2026-10-22) and are compared without case. Microsoft writes some with an en dash where
// the reference shows a hyphen, so dashes are compared as hyphens.
//
// Only successful changes are judged. A failed attempt changed nothing; it stays in the feed.

export const RULES_VERSION = 1;

const iso = (value: Date) => value.toISOString();

export function normalName(value: string | null | undefined): string {
  return (value ?? "")
    .replace(/[‐-―−]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

// Built-in role template IDs from Microsoft's "Microsoft Entra built-in roles" reference
// (https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference,
// checked 2026-10-22). The ID identifies a role even if its display name changes.
export const privilegedRoles: ReadonlyMap<string, string> = new Map([
  ["62e90394-69f5-4237-9190-012177145e10", "Global Administrator"],
  ["e8611ab8-c189-46e8-94e1-60213ab1f814", "Privileged Role Administrator"],
  ["7be44c8a-adaf-4e2a-84d6-ab2649e08a13", "Privileged Authentication Administrator"],
  ["194ae4cb-b126-40b2-bd5b-6091b380977d", "Security Administrator"],
  ["b1be1c3e-b65d-4f19-8427-f6fa0d97feb9", "Conditional Access Administrator"],
  ["29232cdf-9323-42fd-ade2-1d097af3e4de", "Exchange Administrator"],
  ["f28a1f50-f6e7-4571-818b-6a12f2af6b6c", "SharePoint Administrator"],
  ["fe930be7-5e62-47db-91af-98c3a49a38b1", "User Administrator"],
  ["9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3", "Application Administrator"],
  ["158c047a-c907-4556-b7ef-446551a6b5f7", "Cloud Application Administrator"],
  ["c4e39bd9-1100-46d3-8c65-fb160da0071f", "Authentication Administrator"],
  ["3a2c62db-5318-420d-8d74-23affee5d9d5", "Intune Administrator"],
]);
const GLOBAL_ADMINISTRATOR = "Global Administrator";
const roleByName = new Map([...privilegedRoles.values()].map((name) => [normalName(name), name]));

const scores: Record<Extract<Severity, "critical" | "high" | "medium">, number> = {
  critical: 90,
  high: 70,
  medium: 45,
};

const sharedLimitations = [
  "Microsoft publishes directory audit records with a delay and keeps them 7 days without Entra ID P1 or P2, 30 days with it. Prism365 judges only what it collected.",
  "This records a change and who Microsoft says made it. It is not a judgement about that person.",
];

function properties(event: DirectoryChangeEvent, name: string) {
  const wanted = normalName(name);
  return event.targets.flatMap((target) =>
    target.properties.filter((property) => normalName(property.name) === wanted),
  );
}

function newValue(event: DirectoryChangeEvent, name: string): string | null {
  return properties(event, name).find((property) => property.newValue)?.newValue ?? null;
}

const typed = (event: DirectoryChangeEvent, ...types: string[]) =>
  event.targets.find((target) => types.includes(normalName(target.type)));

function actorName(event: DirectoryChangeEvent): string {
  const actor = event.initiatedBy;
  return (
    actor.name ??
    actor.principal ??
    (actor.kind === "app" ? "An application" : "Someone Microsoft did not name")
  );
}

function targetName(target: ChangeTarget | undefined, fallback: string): string {
  return target?.name ?? target?.principal ?? target?.id ?? fallback;
}

type Subject = FindingInput["subject"];

function subjectFor(target: ChangeTarget | undefined, tenantId: string): Subject {
  const label = (target?.name ?? target?.principal)?.slice(0, 200);
  if (target?.id)
    return {
      kind: normalName(target.type) === "user" ? "person" : "resource",
      id: target.id.toLowerCase().slice(0, 200),
      ...(label ? { label } : {}),
    };
  return { kind: "tenant", id: tenantId, label: "This tenant" };
}

interface Draft {
  severity: keyof typeof scores;
  subject: Subject;
  title: string;
  summary: string;
  target?: ChangeTarget | undefined;
  fields?: Record<string, string | number | boolean | null>;
  limitations?: string[];
}

const cut = (value: string | null, max = 500) => (value === null ? null : value.slice(0, max));

function build(rule: ChangeRule, event: DirectoryChangeEvent, draft: Draft): FindingInput {
  return {
    ruleId: rule.id,
    ruleVersion: rule.version,
    dedupeKey: event.id,
    purpose: "security",
    severity: draft.severity,
    score: scores[draft.severity],
    subject: draft.subject,
    title: draft.title,
    summary: draft.summary.slice(0, 1_000),
    evidence: [
      {
        source: "Microsoft Entra audit log",
        ref: event.id,
        observedAt: iso(event.activityAt),
        fields: {
          activity: cut(event.activity),
          category: cut(event.category),
          result: cut(event.result),
          actor: cut(event.initiatedBy.name),
          actorPrincipal: cut(event.initiatedBy.principal),
          actorKind: event.initiatedBy.kind,
          target: cut(draft.target?.name ?? null),
          targetPrincipal: cut(draft.target?.principal ?? null),
          targetType: cut(draft.target?.type ?? null),
          ...draft.fields,
        },
      },
    ],
    limitations: [...sharedLimitations, ...(draft.limitations ?? [])],
    window: { start: event.activityAt, end: event.activityAt },
    observedAt: event.activityAt,
  };
}

export interface ChangeRule {
  id: string;
  version: number;
  // Lower-case category and activity names the rule looks at.
  category: string;
  activities: readonly string[];
  judge(event: DirectoryChangeEvent, tenantId: string): FindingInput | null;
}

function matches(rule: ChangeRule, event: DirectoryChangeEvent): boolean {
  return (
    event.result === "success" &&
    normalName(event.category) === rule.category &&
    rule.activities.includes(normalName(event.activity))
  );
}

// The role a role-assignment record names, by template ID first and then by display name, as
// long as it is one of the privileged roles above.
export function privilegedRole(event: DirectoryChangeEvent): string | null {
  for (const property of properties(event, "Role.TemplateId")) {
    const role = privilegedRoles.get(normalName(property.newValue));
    if (role) return role;
  }
  const names = [
    ...properties(event, "Role.DisplayName").map((property) => property.newValue),
    ...event.targets
      .filter((target) => normalName(target.type) === "role")
      .map((target) => target.name),
  ];
  for (const name of names) {
    const role = roleByName.get(normalName(name));
    if (role) return role;
  }
  return null;
}

const roleAssigned: ChangeRule = {
  id: "role.privileged-assigned",
  version: RULES_VERSION,
  category: "rolemanagement",
  activities: [
    "add member to role",
    "add eligible member to role",
    "add scoped member to role",
    "add member to role scoped over restricted management administrative unit",
  ],
  judge(event, tenantId) {
    if (!matches(this, event)) return null;
    const role = privilegedRole(event);
    if (!role) return null;
    const member =
      typed(event, "user") ??
      typed(event, "serviceprincipal", "group") ??
      event.targets.find((target) => normalName(target.type) !== "role");
    const activity = normalName(event.activity);
    const how = activity.includes("eligible")
      ? " as an eligible assignment"
      : activity.includes("scoped")
        ? ", scoped to an administrative unit"
        : "";
    return build(this, event, {
      severity: role === GLOBAL_ADMINISTRATOR ? "critical" : "high",
      subject: subjectFor(member, tenantId),
      target: member,
      title: `${role} assigned`,
      summary: `${actorName(event)} added ${targetName(member, "an account")} to ${role}${how}. Check that the assignment was requested and is still needed.`,
      fields: { role },
      limitations: [
        "Activating an eligible role through Privileged Identity Management is recorded as a role addition too.",
        "Roles granted through membership of a role-assignable group show up as a group change, not here.",
      ],
    });
  },
};

const keyIds = (value: string | null) =>
  new Set((value ?? "").match(/KeyIdentifier=[0-9A-Fa-f-]{8,}/g)?.map((id) => id.toLowerCase()));

// Whether a credentials record added one. Microsoft lists the application's keys before and
// after in KeyDescription; a key ID only in the newer list was added. When the values are
// missing or were cut, the record cannot be compared, so it counts as an addition.
export function credentialAdded(event: DirectoryChangeEvent): "added" | "unclear" | "removed" {
  if (normalName(event.activity) === "add service principal credentials") return "added";
  const keys = properties(event, "KeyDescription");
  if (keys.length === 0) return "unclear";
  for (const property of keys) {
    const before = keyIds(property.oldValue);
    const after = keyIds(property.newValue);
    if ([...after].some((id) => !before.has(id))) return "added";
    if (property.oldCut || property.newCut) return "unclear";
    if (before.size === 0 && after.size === 0 && (property.oldValue || property.newValue))
      return "unclear";
  }
  return "removed";
}

const credentialAddedRule: ChangeRule = {
  id: "app.credential-added",
  version: RULES_VERSION,
  category: "applicationmanagement",
  activities: [
    "add service principal credentials",
    "update application - certificates and secrets management",
  ],
  judge(event, tenantId) {
    if (!matches(this, event)) return null;
    const outcome = credentialAdded(event);
    if (outcome === "removed") return null;
    const app = typed(event, "application", "serviceprincipal") ?? event.targets[0];
    const name = targetName(app, "an application");
    return build(this, event, {
      severity: "high",
      subject: subjectFor(app, tenantId),
      target: app,
      title: "Credential added to an application",
      summary:
        outcome === "added"
          ? `${actorName(event)} added a certificate or secret to ${name}. Whoever holds it can sign in as the application and use its permissions.`
          : `${actorName(event)} changed the certificates or secrets of ${name}. The record does not show clearly whether one was added, so it is flagged.`,
      fields: { credentialChange: outcome },
      limitations: [
        "Prism365 compares the key IDs Microsoft lists before and after the change; it never sees the secret or certificate itself.",
      ],
    });
  },
};

const isTrue = (value: string | null) => normalName(value) === "true";

const consentRule: ChangeRule = {
  id: "app.consent-granted",
  version: RULES_VERSION,
  category: "applicationmanagement",
  activities: ["consent to application"],
  judge(event, tenantId) {
    if (!matches(this, event)) return null;
    const app = typed(event, "serviceprincipal", "application") ?? event.targets[0];
    const forAll = newValue(event, "ConsentContext.OnBehalfOfAll");
    const admin = isTrue(newValue(event, "ConsentContext.IsAdminConsent"));
    // OnBehalfOfAll says whether consent covered every user. Older records may lack it; then an
    // admin consent is taken to cover the tenant, because that is what the admin center grants.
    const tenantWide = forAll === null ? admin : isTrue(forAll);
    const name = targetName(app, "an application");
    return build(this, event, {
      severity: tenantWide ? "high" : "medium",
      subject: subjectFor(app, tenantId),
      target: app,
      title: tenantWide
        ? "Admin consent granted for the whole tenant"
        : "Consent granted to an application",
      summary: tenantWide
        ? `${actorName(event)} granted ${name} consent on behalf of everyone in the tenant. Check which permissions it received.`
        : `${actorName(event)} consented to ${name}. Check which permissions it received.`,
      fields: {
        tenantWide,
        adminConsent: admin,
        permissions: cut(newValue(event, "ConsentAction.Permissions")),
      },
      limitations: [
        "Whether consent covered the whole tenant is read from the ConsentContext values in Microsoft's record.",
      ],
    });
  },
};

const conditionalAccessRule: ChangeRule = {
  id: "policy.conditional-access-changed",
  version: RULES_VERSION,
  category: "policy",
  activities: [
    "add conditional access policy",
    "update conditional access policy",
    "delete conditional access policy",
  ],
  judge(event, tenantId) {
    if (!matches(this, event)) return null;
    const policy = typed(event, "policy") ?? event.targets[0];
    const verb = normalName(event.activity).split(" ")[0] ?? "update";
    const done = { add: "added", update: "updated", delete: "deleted" }[verb] ?? "changed";
    return build(this, event, {
      severity: verb === "delete" ? "high" : "medium",
      subject: subjectFor(policy, tenantId),
      target: policy,
      title: `Conditional Access policy ${done}`,
      summary: `${actorName(event)} ${done} the Conditional Access policy ${targetName(policy, "with no name in the record")}. Check that the change was planned.`,
      limitations: [
        "The record holds the policy before and after, cut to 400 characters each; open the policy in the Entra admin center for every setting.",
      ],
    });
  },
};

const federationRule: ChangeRule = {
  id: "domain.federation-changed",
  version: RULES_VERSION,
  category: "directorymanagement",
  activities: ["set domain authentication", "set federation settings on domain"],
  judge(event, tenantId) {
    if (!matches(this, event)) return null;
    const domain = typed(event, "domain") ?? event.targets[0];
    return build(this, event, {
      severity: "critical",
      subject: { kind: "tenant", id: tenantId, label: "This tenant" },
      target: domain,
      title: "Domain sign-in settings changed",
      summary: `${actorName(event)} changed how ${targetName(domain, "a domain")} signs people in (${event.activity ?? "domain authentication"}). A federated domain trusts another identity provider to sign its users in, so an unplanned change can hand sign-in to someone else.`,
      fields: { domain: cut(domain?.name ?? null) },
      limitations: [
        "The record says the setting changed; it does not show the identity provider's certificate or address.",
      ],
    });
  },
};

const securityInfoRule: ChangeRule = {
  id: "auth-method.removed-by-admin",
  version: RULES_VERSION,
  category: "usermanagement",
  activities: ["admin deleted security info"],
  judge(event, tenantId) {
    if (!matches(this, event)) return null;
    const person = typed(event, "user") ?? event.targets[0];
    // Removing your own security info is not removing someone else's.
    if (
      event.initiatedBy.kind === "user" &&
      event.initiatedBy.id &&
      person?.id &&
      event.initiatedBy.id.toLowerCase() === person.id.toLowerCase()
    )
      return null;
    return build(this, event, {
      severity: "medium",
      subject: subjectFor(person, tenantId),
      target: person,
      title: "Security info removed by an administrator",
      summary: `${actorName(event)} deleted a sign-in method of ${targetName(person, "an account")}. If the person did not ask for it, check who registered methods on the account afterwards.`,
      limitations: ["Microsoft's record does not always say which method was removed."],
    });
  },
};

export const changeRules: readonly ChangeRule[] = [
  roleAssigned,
  credentialAddedRule,
  consentRule,
  conditionalAccessRule,
  federationRule,
  securityInfoRule,
];

// Every finding one stored change produces.
export function judgeChange(event: DirectoryChangeEvent, tenantId: string): FindingInput[] {
  return changeRules
    .map((rule) => rule.judge(event, tenantId))
    .filter((finding): finding is FindingInput => finding !== null);
}
