import type { Severity } from "@prism365/contracts";
import { privilegedRoles } from "@prism365/graph";
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

// The privileged roles live in @prism365/graph (built-in role template IDs from Microsoft's role
// reference, checked 2026-10-22), where the audits read them too. Re-exported so this module keeps
// its public surface.
export { privilegedRoles };
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
  // One finding per event unless the rule says what counts as the same thing.
  dedupeKey?: string;
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
    dedupeKey: (draft.dedupeKey ?? event.id).slice(0, 200),
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
  judge(event: DirectoryChangeEvent, tenantId: string, related: Related): FindingInput | null;
}

// What a rule may know about other stored changes: the correlation ids of the PIM activation
// records stored so far among those the event could belong to.
export interface Related {
  activationCorrelations: ReadonlySet<string>;
}

const nothingRelated: Related = { activationCorrelations: new Set() };

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

export const PIM_ACTIVATION = "add member to role completed (pim activation)";

// A role addition the directory records on behalf of Privileged Identity Management: "Add member
// to role" made by Microsoft's PIM service principal, named "MS-PIM" in observed tenants (not in
// Learn's reference). PIM writes these for activations and for permanent assignments alike.
export function addedByPim(event: DirectoryChangeEvent): boolean {
  return (
    normalName(event.activity) === "add member to role" &&
    event.initiatedBy.kind === "app" &&
    /^ms-pim/.test(normalName(event.initiatedBy.name))
  );
}

// An eligible role made active. Only the PIM service's own "Add member to role completed (PIM
// activation)" record says so; the directory's record of the same activation counts only when
// it shares that record's correlation id. Anything else, a permanent assignment made through PIM
// included, is an assignment: one finding per event, so it can never fold into an activation
// someone has already acknowledged.
export function isActivation(
  event: DirectoryChangeEvent,
  related: Related = nothingRelated,
): boolean {
  if (normalName(event.activity) === PIM_ACTIVATION) return true;
  return (
    addedByPim(event) &&
    event.correlationId !== null &&
    related.activationCorrelations.has(event.correlationId)
  );
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
    "add member to role in pim completed (permanent)",
    "add member to role in pim completed (timebound)",
    "add eligible member to role in pim completed (permanent)",
    "add eligible member to role in pim completed (timebound)",
    PIM_ACTIVATION,
  ],
  judge(event, tenantId, related) {
    if (!matches(this, event)) return null;
    const role = privilegedRole(event);
    if (!role) return null;
    const member =
      typed(event, "user") ??
      typed(event, "serviceprincipal", "group") ??
      event.targets.find((target) => normalName(target.type) !== "role");
    const subject = subjectFor(member, tenantId);
    const severity = role === GLOBAL_ADMINISTRATOR ? "critical" : "high";
    // An eligible administrator may activate several times a day: one finding per person, role
    // and day, rather than one per activation. A permanent assignment stays one per event.
    if (isActivation(event, related))
      return build(this, event, {
        dedupeKey: `activation:${normalName(role)}:${subject.id}:${iso(event.activityAt).slice(0, 10)}`,
        severity,
        subject,
        target: member,
        title: `${role} activated`,
        summary: `${targetName(member, "An account")} activated ${role} through Privileged Identity Management on ${iso(event.activityAt).slice(0, 10)}. Activations of the same role by the same account that day are counted in this one finding.`,
        fields: { role, activation: true },
        limitations: [
          "Each activation is counted in the finding's occurrences; the evidence shows the latest.",
        ],
      });
    const activity = normalName(event.activity);
    const how = activity.includes("eligible")
      ? " as an eligible assignment"
      : activity.includes("scoped")
        ? ", scoped to an administrative unit"
        : activity.includes(" in pim ") || addedByPim(event)
          ? " through Privileged Identity Management"
          : "";
    return build(this, event, {
      severity,
      subject,
      target: member,
      title: `${role} assigned`,
      summary: `${actorName(event)} added ${targetName(member, "an account")} to ${role}${how}. Check that the assignment was requested and is still needed.`,
      fields: { role },
      limitations: [
        "Roles granted through membership of a role-assignable group show up as a group change, not here.",
        ...(addedByPim(event)
          ? [
              "PIM writes this directory record for activations too. It is counted with an activation only when Microsoft gives both records the same correlation id, never by time, so an activation recorded with different ids also raises this finding.",
            ]
          : []),
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

// Learn's reference has no federation-specific "Update domain" activity, only the general one.
// It counts here when what it changed is named as an authentication or federation setting.
const signInSetting = /authentication|federat/;

const federationRule: ChangeRule = {
  id: "domain.federation-changed",
  version: RULES_VERSION,
  category: "directorymanagement",
  activities: ["set domain authentication", "set federation settings on domain", "update domain"],
  judge(event, tenantId) {
    if (!matches(this, event)) return null;
    if (
      normalName(event.activity) === "update domain" &&
      !event.targets.some((target) =>
        target.properties.some((property) => signInSetting.test(normalName(property.name))),
      )
    )
      return null;
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

// An "Update user" that leaves StrongAuthenticationMethod empty after it held methods: an
// administrator made the person register MFA again, which also drops every method they had.
export function methodsCleared(event: DirectoryChangeEvent): boolean {
  return properties(event, "StrongAuthenticationMethod").some(
    (property) =>
      // An old value cut for length held methods too; what counts is an empty new value.
      (property.oldValue ?? "").trim() !== "" &&
      ["", "[]"].includes((property.newValue ?? "").trim()),
  );
}

const securityInfoRule: ChangeRule = {
  id: "auth-method.removed-by-admin",
  version: RULES_VERSION,
  category: "usermanagement",
  activities: ["admin deleted security info", "update user"],
  judge(event, tenantId) {
    if (!matches(this, event)) return null;
    const cleared = normalName(event.activity) === "update user";
    if (cleared && !methodsCleared(event)) return null;
    const person = typed(event, "user") ?? event.targets[0];
    // Removing your own security info is not removing someone else's.
    if (
      event.initiatedBy.kind === "user" &&
      event.initiatedBy.id &&
      person?.id &&
      event.initiatedBy.id.toLowerCase() === person.id.toLowerCase()
    )
      return null;
    if (cleared)
      return build(this, event, {
        severity: "medium",
        subject: subjectFor(person, tenantId),
        target: person,
        title: "Sign-in methods cleared by an administrator",
        summary: `${actorName(event)} cleared the registered sign-in methods of ${targetName(person, "an account")}, so they must register MFA again. If the person did not ask for it, check who registers methods on the account next.`,
        limitations: [
          "Read from the StrongAuthenticationMethod value in an Update user record; methods registered as security info may be listed elsewhere.",
        ],
      });
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
export function judgeChange(
  event: DirectoryChangeEvent,
  tenantId: string,
  related: Related = nothingRelated,
): FindingInput[] {
  return changeRules
    .map((rule) => rule.judge(event, tenantId, related))
    .filter((finding): finding is FindingInput => finding !== null);
}
