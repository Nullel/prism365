import { checkFinding } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  type ChangeTarget,
  type DirectoryChangeEvent,
  judgeChange,
  privilegedRoles,
} from "../src/index.ts";

// Synthetic events only: invented people, apps and domains.

const tenantId = "0a1b2c3d-0000-4000-8000-0000000000aa";
const adminId = "0a1b2c3d-0000-4000-8000-00000000a001";
const personId = "0a1b2c3d-0000-4000-8000-00000000b002";
const globalAdministrator = "62e90394-69f5-4237-9190-012177145e10";

const person: ChangeTarget = {
  id: personId,
  name: "Synthetic Person",
  type: "User",
  principal: "person@example.invalid",
  properties: [],
};

function change(overrides: Partial<DirectoryChangeEvent> = {}): DirectoryChangeEvent {
  return {
    id: "Directory_rule_1",
    activityAt: new Date("2026-10-21T09:15:02.721Z"),
    activity: "Update user",
    category: "UserManagement",
    operationType: "Update",
    result: "success",
    resultReason: null,
    loggedByService: "Core Directory",
    correlationId: null,
    initiatedBy: {
      kind: "user",
      id: adminId,
      name: "Synthetic Admin",
      principal: "admin@example.invalid",
    },
    targets: [person],
    details: [],
    ...overrides,
  };
}

const property = (name: string, newValue: string | null, oldValue: string | null = null) => ({
  name,
  oldValue,
  newValue,
});

const roleAdded = (role: { id?: string; name?: string }, activity = "Add member to role") =>
  change({
    activity,
    category: "RoleManagement",
    targets: [
      {
        ...person,
        properties: [
          ...(role.name ? [property("Role.DisplayName", role.name)] : []),
          ...(role.id ? [property("Role.TemplateId", role.id)] : []),
        ],
      },
    ],
  });

// Every finding a rule produces must pass the same checks the store applies.
function judged(event: DirectoryChangeEvent) {
  const findings = judgeChange(event, tenantId);
  for (const finding of findings) expect(() => checkFinding(finding)).not.toThrow();
  return findings;
}

describe("role.privileged-assigned", () => {
  it("flags Global Administrator as critical, by template ID, about the person added", () => {
    const [finding] = judged(roleAdded({ id: globalAdministrator }));
    expect(finding).toMatchObject({
      ruleId: "role.privileged-assigned",
      ruleVersion: 1,
      dedupeKey: "Directory_rule_1",
      purpose: "security",
      severity: "critical",
      subject: { kind: "person", id: personId, label: "Synthetic Person" },
      title: "Global Administrator assigned",
    });
    expect(finding?.summary).toBe(
      "Synthetic Admin added Synthetic Person to Global Administrator. Check that the assignment was requested and is still needed.",
    );
    expect(finding?.evidence[0]).toMatchObject({
      ref: "Directory_rule_1",
      observedAt: "2026-10-21T09:15:02.721Z",
      fields: {
        activity: "Add member to role",
        actor: "Synthetic Admin",
        target: "Synthetic Person",
        role: "Global Administrator",
      },
    });
  });

  it("flags the other privileged roles as high, by name and without case", () => {
    expect(privilegedRoles.size).toBe(12);
    for (const name of privilegedRoles.values()) {
      if (name === "Global Administrator") continue;
      const [finding] = judged(
        roleAdded({ name: name.toUpperCase() }, "ADD ELIGIBLE MEMBER TO ROLE"),
      );
      expect(finding?.severity, name).toBe("high");
      expect(finding?.summary).toContain("as an eligible assignment");
    }
    // The role can also be named by a target of type Role.
    const byTarget = change({
      activity: "Add member to role",
      category: "rolemanagement",
      targets: [
        person,
        {
          id: "role-1",
          name: "Security Administrator",
          type: "Role",
          principal: null,
          properties: [],
        },
      ],
    });
    expect(judged(byTarget)[0]?.title).toBe("Security Administrator assigned");
  });

  it("ignores a role that is not privileged, a removal and a failed attempt", () => {
    expect(judged(roleAdded({ name: "Reports Reader" }))).toEqual([]);
    expect(
      judged({ ...roleAdded({ id: globalAdministrator }), activity: "Remove member from role" }),
    ).toEqual([]);
    expect(judged({ ...roleAdded({ id: globalAdministrator }), result: "failure" })).toEqual([]);
    expect(
      judged({ ...roleAdded({ id: globalAdministrator }), category: "GroupManagement" }),
    ).toEqual([]);
  });
});

describe("app.credential-added", () => {
  const app: ChangeTarget = {
    id: "app-object-1",
    name: "Synthetic Payroll Connector",
    type: "Application",
    principal: null,
    properties: [],
  };
  const keys = (...ids: string[]) =>
    ids
      .map((id) => `[KeyIdentifier=${id}, KeyType=Password, KeyUsage=Verify, DisplayName=s]`)
      .join(", ");
  const updated = (oldValue: string, newValue: string, cut = false) =>
    change({
      activity: "Update application – Certificates and secrets management ",
      category: "ApplicationManagement",
      targets: [
        {
          ...app,
          properties: [
            {
              name: "KeyDescription",
              oldValue,
              newValue,
              ...(cut ? { newCut: true as const } : {}),
            },
          ],
        },
      ],
    });
  const one = "0a1b2c3d-1111-4000-8000-000000000001";
  const two = "0a1b2c3d-2222-4000-8000-000000000002";

  it("flags a new key on an application, high, about the application", () => {
    const [finding] = judged(updated(keys(one), keys(one, two)));
    expect(finding).toMatchObject({
      ruleId: "app.credential-added",
      severity: "high",
      subject: { kind: "resource", id: "app-object-1", label: "Synthetic Payroll Connector" },
    });
    expect(finding?.summary).toContain("added a certificate or secret");
  });

  it("flags service principal credentials, and a record too long to compare", () => {
    expect(
      judged(
        change({
          activity: "Add service principal credentials",
          category: "ApplicationManagement",
          targets: [{ ...app, type: "ServicePrincipal" }],
        }),
      ),
    ).toHaveLength(1);
    const [unclear] = judged(updated(keys(one, two), keys(one), true));
    expect(unclear?.summary).toContain("does not show clearly");
  });

  it("ignores a key that was only removed, and a failed attempt", () => {
    expect(judged(updated(keys(one, two), keys(one)))).toEqual([]);
    expect(judged({ ...updated(keys(one), keys(one, two)), result: "failure" })).toEqual([]);
  });
});

describe("app.consent-granted", () => {
  const app: ChangeTarget = {
    id: "sp-1",
    name: "Synthetic Survey App",
    type: "ServicePrincipal",
    principal: null,
    properties: [],
  };
  const consent = (admin: string, all: string | null) =>
    change({
      activity: "Consent to application",
      category: "ApplicationManagement",
      targets: [
        {
          ...app,
          properties: [
            property("ConsentContext.IsAdminConsent", admin),
            ...(all === null ? [] : [property("ConsentContext.OnBehalfOfAll", all)]),
            property("ConsentAction.Permissions", "Scope: User.Read Mail.Read"),
          ],
        },
      ],
    });

  it("rates admin consent for the whole tenant high and one user's consent medium", () => {
    const [tenantWide] = judged(consent("True", "True"));
    expect(tenantWide).toMatchObject({
      severity: "high",
      subject: { kind: "resource", id: "sp-1" },
      title: "Admin consent granted for the whole tenant",
    });
    expect(tenantWide?.evidence[0]?.fields.permissions).toBe("Scope: User.Read Mail.Read");
    expect(judged(consent("False", "False"))[0]?.severity).toBe("medium");
    // A record without OnBehalfOfAll falls back on IsAdminConsent.
    expect(judged(consent("True", null))[0]?.severity).toBe("high");
  });

  it("ignores the permission grants that come with consent, and a failed consent", () => {
    expect(
      judged({ ...consent("True", "True"), activity: "Add delegated permission grant" }),
    ).toEqual([]);
    expect(judged({ ...consent("True", "True"), result: "failure" })).toEqual([]);
  });
});

describe("policy.conditional-access-changed", () => {
  const policy = (activity: string) =>
    change({
      activity,
      category: "Policy",
      targets: [
        {
          id: "policy-1",
          name: "Require MFA for admins",
          type: "Policy",
          principal: null,
          properties: [],
        },
      ],
    });

  it("rates an update medium and a deletion high, about the policy", () => {
    expect(judged(policy("Update Conditional Access policy"))[0]).toMatchObject({
      severity: "medium",
      title: "Conditional Access policy updated",
      subject: { kind: "resource", id: "policy-1", label: "Require MFA for admins" },
    });
    expect(judged(policy("Add Conditional Access policy"))[0]?.title).toBe(
      "Conditional Access policy added",
    );
    expect(judged(policy("Delete Conditional Access policy"))[0]?.severity).toBe("high");
  });

  it("ignores named locations", () => {
    expect(judged(policy("Update named location"))).toEqual([]);
  });
});

describe("domain.federation-changed", () => {
  const domain = (activity: string) =>
    change({
      activity,
      category: "DirectoryManagement",
      targets: [
        {
          id: "domain-1",
          name: "larkspur.example",
          type: "Domain",
          principal: null,
          properties: [],
        },
      ],
    });

  it("flags federation settings as critical, about the tenant", () => {
    for (const activity of ["Set federation settings on domain", "Set domain authentication"])
      expect(judged(domain(activity))[0]).toMatchObject({
        severity: "critical",
        subject: { kind: "tenant", id: tenantId },
        evidence: [{ fields: { domain: "larkspur.example" } }],
      });
  });

  it("ignores verifying a domain", () => {
    expect(judged(domain("Verify domain"))).toEqual([]);
  });
});

describe("auth-method.removed-by-admin", () => {
  const deleted = (initiator: string) =>
    change({
      activity: "Admin deleted security info",
      category: "UserManagement",
      initiatedBy: { kind: "user", id: initiator, name: "Synthetic Admin", principal: null },
    });

  it("flags an administrator removing someone else's method, medium, about the person", () => {
    expect(judged(deleted(adminId))[0]).toMatchObject({
      ruleId: "auth-method.removed-by-admin",
      severity: "medium",
      subject: { kind: "person", id: personId },
    });
  });

  it("ignores removing your own, and a person removing their own method", () => {
    expect(judged(deleted(personId.toUpperCase()))).toEqual([]);
    expect(judged({ ...deleted(adminId), activity: "User deleted security info" })).toEqual([]);
  });
});

describe("every rule", () => {
  it("judges an ordinary change as nothing", () => {
    expect(judged(change())).toEqual([]);
  });

  it("names nobody it does not know, and still passes the finding checks", () => {
    const [finding] = judged({
      ...roleAdded({ id: globalAdministrator }),
      initiatedBy: { kind: "unknown", id: null, name: null, principal: null },
      targets: [
        {
          id: null,
          name: null,
          type: "User",
          principal: null,
          properties: [property("Role.TemplateId", globalAdministrator)],
        },
      ],
    });
    expect(finding?.summary).toContain("Someone Microsoft did not name added an account");
    expect(finding?.subject).toEqual({ kind: "tenant", id: tenantId, label: "This tenant" });
  });
});

describe("rule activities the first tests left out", () => {
  it("flags scoped and restricted administrative unit role additions", () => {
    const scoped = judged(roleAdded({ name: "User Administrator" }, "Add scoped member to role"));
    expect(scoped[0]?.summary).toContain("scoped to an administrative unit");
    expect(
      judged(
        roleAdded(
          { name: "User Administrator" },
          "Add member to role scoped over Restricted Management Administrative Unit",
        ),
      ),
    ).toHaveLength(1);
  });

  it("takes an admin consent that did not cover everyone as one user's consent", () => {
    const [finding] = judged(
      change({
        activity: "Consent to application",
        category: "ApplicationManagement",
        targets: [
          {
            id: "sp-1",
            name: "Synthetic Survey App",
            type: "ServicePrincipal",
            principal: null,
            properties: [
              property("ConsentContext.IsAdminConsent", "True"),
              property("ConsentContext.OnBehalfOfAll", "False"),
            ],
          },
        ],
      }),
    );
    expect(finding?.severity).toBe("medium");
  });

  it("flags a credentials record with no key list, as unclear", () => {
    const [finding] = judged(
      change({
        activity: "Update application - Certificates and secrets management",
        category: "ApplicationManagement",
        targets: [
          {
            id: "app-1",
            name: "Synthetic App",
            type: "Application",
            principal: null,
            properties: [],
          },
        ],
      }),
    );
    expect(finding?.evidence[0]?.fields.credentialChange).toBe("unclear");
  });
});

describe("role activations through Privileged Identity Management", () => {
  const activation = (at: string, how: "pim" | "directory" = "pim") =>
    change({
      id: `Directory_${at}_${how}`,
      activityAt: new Date(at),
      activity:
        how === "pim" ? "Add member to role completed (PIM activation)" : "Add member to role",
      category: "RoleManagement",
      initiatedBy:
        how === "pim"
          ? { kind: "user", id: personId, name: "Synthetic Person", principal: null }
          : { kind: "app", id: "sp-pim", name: "MS-PIM", principal: null },
      targets: [
        {
          ...person,
          properties: [property("Role.TemplateId", globalAdministrator)],
        },
      ],
    });

  it("counts one person's activations of a role in one finding a day", () => {
    const [morning] = judged(activation("2026-10-21T08:00:00.000Z"));
    const [afternoon] = judged(activation("2026-10-21T15:00:00.000Z", "directory"));
    const [nextDay] = judged(activation("2026-10-22T08:00:00.000Z"));
    expect(morning).toMatchObject({
      severity: "critical",
      title: "Global Administrator activated",
    });
    expect(morning?.dedupeKey).toBe(afternoon?.dedupeKey);
    expect(nextDay?.dedupeKey).not.toBe(morning?.dedupeKey);
    // A permanent assignment stays one finding per event.
    expect(judged(roleAdded({ id: globalAdministrator }))[0]?.dedupeKey).toBe("Directory_rule_1");
  });
});

describe("an administrator forcing MFA registration again", () => {
  const cleared = (oldValue: string, newValue: string, initiator = adminId) =>
    change({
      activity: "Update user",
      category: "UserManagement",
      initiatedBy: { kind: "user", id: initiator, name: "Synthetic Admin", principal: null },
      targets: [
        {
          ...person,
          properties: [property("StrongAuthenticationMethod", newValue, oldValue)],
        },
      ],
    });
  const methods = '[{"MethodType":6,"Default":true}]';

  it("flags clearing someone else's sign-in methods", () => {
    const [finding] = judged(cleared(methods, ""));
    expect(finding).toMatchObject({
      ruleId: "auth-method.removed-by-admin",
      severity: "medium",
      title: "Sign-in methods cleared by an administrator",
      subject: { kind: "person", id: personId },
    });
  });

  it("ignores other updates, a list that was already empty, and your own account", () => {
    expect(judged(cleared(methods, '[{"MethodType":7,"Default":true}]'))).toEqual([]);
    expect(judged(cleared("", ""))).toEqual([]);
    expect(judged(cleared(methods, "", personId))).toEqual([]);
    expect(judged(change())).toEqual([]);
  });
});

describe("a domain update that changes how it signs people in", () => {
  const update = (name: string) =>
    change({
      activity: "Update domain",
      category: "DirectoryManagement",
      targets: [
        {
          id: "domain-1",
          name: "larkspur.example",
          type: "Domain",
          principal: null,
          properties: [property(name, "Federated", "Managed")],
        },
      ],
    });

  it("flags an authentication or federation setting, and nothing else", () => {
    expect(judged(update("AuthenticationType"))[0]?.severity).toBe("critical");
    expect(judged(update("FederationBrandName"))).toHaveLength(1);
    expect(judged(update("IsDefault"))).toEqual([]);
  });
});
