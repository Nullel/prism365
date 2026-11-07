import { describe, expect, it } from "vitest";
import { emptyPerson, type Person } from "../src/combined-risks-inputs.ts";
import { type MfaPolicy, mfaPolicies, mfaReach } from "../src/combined-risks-reach.ts";
import { policySchema } from "../src/conditional-access.ts";
import { RULES, type RuleEnv, type RuleId } from "../src/combined-risks-rules.ts";
import { daysAgo, finding, now } from "./combined-risks-support.ts";

// Each rule on its own, both ways, over a joined person. Synthetic data only.

const rule = (id: RuleId) => {
  const found = RULES.find((item) => item.id === id);
  if (!found) throw new Error(`No rule ${id}`);
  return found;
};

const person = (over: Partial<Person> = {}): Person => ({ ...emptyPerson("u-1"), ...over });
const env: RuleEnv = { now, policies: null };
const check = (id: RuleId, who: Person, over: Partial<RuleEnv> = {}) =>
  rule(id).check(who, { ...env, ...over });

const admin = (held = "Permanent") => [
  {
    role: "Global Administrator",
    privileged: true,
    held,
    templateId: "62e90394-69f5-4237-9190-012177145e10",
  },
];
const helpdesk = [
  { role: "Helpdesk Administrator", privileged: false, held: "Permanent", templateId: null },
];

// A flat description of the "Users" part of an MFA policy, built into the real parsed policy so
// the tests exercise the same shape the audit reads from Microsoft.
interface PolicySpec {
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
const guests = { guestOrExternalUserTypes: "b2bCollaborationGuest" };
const policy = (over: Partial<PolicySpec> = {}): MfaPolicy => {
  const spec: PolicySpec = {
    id: "p1",
    name: "Require MFA for everyone",
    includeUsers: ["all"],
    excludeUsers: [],
    includeGroups: [],
    excludeGroups: [],
    includeRoles: [],
    excludeRoles: [],
    includesGuests: false,
    excludesGuests: false,
    ...over,
  };
  const parsed = policySchema.parse({
    id: spec.id,
    displayName: spec.name,
    state: "enabled",
    conditions: {
      users: {
        includeUsers: spec.includeUsers,
        excludeUsers: spec.excludeUsers,
        includeGroups: spec.includeGroups,
        excludeGroups: spec.excludeGroups,
        includeRoles: spec.includeRoles,
        excludeRoles: spec.excludeRoles,
        includeGuestsOrExternalUsers: spec.includesGuests ? guests : null,
        excludeGuestsOrExternalUsers: spec.excludesGuests ? guests : null,
      },
    },
    grantControls: { operator: "OR", builtInControls: ["mfa"] },
  });
  return { id: spec.id, name: spec.name, policy: parsed };
};

describe("the rule list", () => {
  it("has nine rules with the agreed severities", () => {
    expect(RULES.map((item) => [item.id, item.severity])).toEqual([
      ["admin-unprotected", "critical"],
      ["risky-unprotected", "critical"],
      ["admin-signin-finding", "critical"],
      ["dormant-admin", "high"],
      ["guest-with-power", "high"],
      ["weak-password-no-mfa", "high"],
      ["admin-weak-device", "high"],
      ["disabled-still-holding", "medium"],
      ["paying-for-nothing", "medium"],
    ]);
  });

  it("gives every rule a plain label without dashes or marks", () => {
    for (const item of RULES) expect(item.label, item.id).toMatch(/^[A-Z][A-Za-z ,-]+$/);
  });
});

describe("admin-unprotected", () => {
  it("trips for a privileged role without MFA, active or eligible, and names the facts", () => {
    for (const held of ["Permanent", "Eligible"]) {
      const trip = check("admin-unprotected", person({ roles: admin(held), mfa: false }));
      expect(trip?.reason).toBe(
        `Holds Global Administrator (${held.toLowerCase()}) and has no MFA registered.`,
      );
    }
  });

  it("stays quiet with MFA registered, an unknown MFA state, or no privileged role", () => {
    expect(check("admin-unprotected", person({ roles: admin(), mfa: true }))).toBeNull();
    expect(check("admin-unprotected", person({ roles: admin(), mfa: null }))).toBeNull();
    expect(check("admin-unprotected", person({ roles: helpdesk, mfa: false }))).toBeNull();
    expect(check("admin-unprotected", person({ mfa: false }))).toBeNull();
  });

  it("trips for an administrator left out of an MFA policy even with MFA registered", () => {
    const left = policy({ excludeRoles: ["62e90394-69f5-4237-9190-012177145e10"] });
    const trip = check("admin-unprotected", person({ roles: admin(), mfa: true }), {
      policies: [left],
    });
    expect(trip).toMatchObject({ usedPolicies: true });
    expect(trip?.reason).toContain("is left out of Require MFA for everyone, which requires MFA");
  });

  it("joins both facts in one sentence", () => {
    const left = policy({ excludeUsers: ["u-1"] });
    const trip = check("admin-unprotected", person({ roles: admin(), mfa: false }), {
      policies: [left],
    });
    expect(trip?.reason).toBe(
      "Holds Global Administrator (permanent) and has no MFA registered and is left out of Require MFA for everyone, which requires MFA.",
    );
  });

  it("does not mention policies that cover the administrator", () => {
    expect(
      check("admin-unprotected", person({ roles: admin(), mfa: true }), { policies: [policy()] }),
    ).toBeNull();
  });

  it("shortens a long role list", () => {
    const many = ["A", "B", "C", "D"].map((role) => ({
      role: `${role} Administrator`,
      privileged: true,
      held: "Permanent",
      templateId: null,
    }));
    expect(check("admin-unprotected", person({ roles: many, mfa: false }))?.reason).toContain(
      "A Administrator (permanent), B Administrator (permanent) and 2 more",
    );
  });
});

describe("risky-unprotected", () => {
  it("trips for high or medium risk that is still open and no MFA", () => {
    for (const level of ["High", "Medium"])
      for (const state of ["At risk", "Confirmed compromised"])
        expect(
          check("risky-unprotected", person({ risk: { level, state }, mfa: false }))?.reason,
        ).toBe(
          `Is marked ${level.toLowerCase()} risk (${state.toLowerCase()}) and has no MFA registered.`,
        );
  });

  it("stays quiet for low risk, closed risk, MFA registered or no risk record", () => {
    const base = { level: "High", state: "At risk" };
    expect(
      check("risky-unprotected", person({ risk: { ...base, level: "Low" }, mfa: false })),
    ).toBeNull();
    expect(
      check("risky-unprotected", person({ risk: { ...base, level: "Hidden" }, mfa: false })),
    ).toBeNull();
    for (const state of ["Remediated", "Dismissed"])
      expect(
        check("risky-unprotected", person({ risk: { ...base, state }, mfa: false })),
      ).toBeNull();
    expect(check("risky-unprotected", person({ risk: base, mfa: true }))).toBeNull();
    expect(check("risky-unprotected", person({ mfa: false }))).toBeNull();
  });

  it("trips for an at-risk account left out of an MFA policy", () => {
    const left = policy({ excludeUsers: ["u-1"] });
    expect(
      check("risky-unprotected", person({ risk: { level: "High", state: "At risk" }, mfa: true }), {
        policies: [left],
      })?.usedPolicies,
    ).toBe(true);
  });
});

describe("admin-signin-finding", () => {
  it("trips for a medium or worse open finding on a privileged account, naming the worst", () => {
    const who = person({
      roles: admin(),
      findings: [
        finding({ severity: "medium", title: "Burst" }),
        finding({ severity: "critical", title: "Spray" }),
      ],
    });
    expect(check("admin-signin-finding", who)?.reason).toBe(
      "Has an open critical severity sign-in finding (Spray) and holds Global Administrator (permanent).",
    );
    expect(
      check(
        "admin-signin-finding",
        person({ roles: admin(), findings: [finding({ severity: "medium" })] }),
      ),
    ).not.toBeNull();
  });

  it("stays quiet for low findings, no findings, or an account without a privileged role", () => {
    expect(
      check(
        "admin-signin-finding",
        person({ roles: admin(), findings: [finding({ severity: "low" })] }),
      ),
    ).toBeNull();
    expect(
      check(
        "admin-signin-finding",
        person({ roles: admin(), findings: [finding({ severity: "info" })] }),
      ),
    ).toBeNull();
    expect(check("admin-signin-finding", person({ roles: admin() }))).toBeNull();
    expect(
      check("admin-signin-finding", person({ roles: helpdesk, findings: [finding()] })),
    ).toBeNull();
    expect(check("admin-signin-finding", person({ findings: [finding()] }))).toBeNull();
  });
});

describe("dormant-admin", () => {
  it("trips at 30 days and not at 29", () => {
    expect(
      check("dormant-admin", person({ roles: admin(), lastSuccess: daysAgo(30) }))?.reason,
    ).toBe("Holds Global Administrator (permanent) and has had no successful sign-in for 30 days.");
    expect(check("dormant-admin", person({ roles: admin(), lastSuccess: daysAgo(29) }))).toBeNull();
  });

  it("trips for no sign-in on record once the account is old enough to count", () => {
    expect(
      check("dormant-admin", person({ roles: admin(), lastSuccess: null, days: 31 })),
    ).not.toBeNull();
    expect(
      check("dormant-admin", person({ roles: admin(), lastSuccess: null, days: 30 })),
    ).not.toBeNull();
    expect(
      check("dormant-admin", person({ roles: admin(), lastSuccess: null, days: 29 })),
    ).toBeNull();
    expect(
      check("dormant-admin", person({ roles: admin(), lastSuccess: null, days: null })),
    ).toBeNull();
  });

  it("needs a privileged role", () => {
    expect(
      check("dormant-admin", person({ roles: helpdesk, lastSuccess: daysAgo(99) })),
    ).toBeNull();
    expect(check("dormant-admin", person({ lastSuccess: daysAgo(99) }))).toBeNull();
  });
});

describe("guest-with-power", () => {
  it("trips for a guest holding any directory role, privileged or not", () => {
    expect(check("guest-with-power", person({ guest: true, roles: helpdesk }))?.reason).toBe(
      "Is a guest account that holds Helpdesk Administrator (permanent).",
    );
    expect(
      check("guest-with-power", person({ guest: true, roles: admin("Eligible") })),
    ).not.toBeNull();
  });

  it("trips for a guest that owns a group, once owners are known", () => {
    expect(
      check("guest-with-power", person({ guest: true, ownedGroups: ["g1", "g2"] }))?.reason,
    ).toBe("Is a guest account that owns 2 groups.");
    expect(
      check("guest-with-power", person({ guest: true, roles: helpdesk, ownedGroups: ["g1"] }))
        ?.reason,
    ).toBe("Is a guest account that holds Helpdesk Administrator (permanent) and owns 1 group.");
  });

  it("stays quiet for a guest with nothing, and for a member with a role", () => {
    expect(check("guest-with-power", person({ guest: true }))).toBeNull();
    expect(check("guest-with-power", person({ guest: false, roles: admin() }))).toBeNull();
  });
});

describe("weak-password-no-mfa", () => {
  const base = { neverExpires: false, weak: false, days: 10 };
  it("trips for each weakness on its own, and says which", () => {
    expect(
      check(
        "weak-password-no-mfa",
        person({ mfa: false, password: { ...base, neverExpires: true } }),
      )?.reason,
    ).toBe("Its password never expires, and no MFA is registered.");
    expect(
      check("weak-password-no-mfa", person({ mfa: false, password: { ...base, weak: true } }))
        ?.reason,
    ).toBe("A strong password is not required, and no MFA is registered.");
    expect(
      check("weak-password-no-mfa", person({ mfa: false, password: { ...base, days: 366 } }))
        ?.reason,
    ).toBe("Its password is 366 days old, and no MFA is registered.");
  });

  it("combines weaknesses in one sentence", () => {
    expect(
      check(
        "weak-password-no-mfa",
        person({ mfa: false, password: { neverExpires: true, weak: true, days: 400 } }),
      )?.reason,
    ).toBe(
      "Its password never expires, a strong password is not required, its password is 400 days old, and no MFA is registered.",
    );
  });

  it("stays quiet at exactly 365 days, with MFA, with an unknown MFA state, or no password row", () => {
    expect(
      check("weak-password-no-mfa", person({ mfa: false, password: { ...base, days: 365 } })),
    ).toBeNull();
    expect(
      check("weak-password-no-mfa", person({ mfa: true, password: { ...base, weak: true } })),
    ).toBeNull();
    expect(
      check("weak-password-no-mfa", person({ mfa: null, password: { ...base, weak: true } })),
    ).toBeNull();
    expect(check("weak-password-no-mfa", person({ mfa: false }))).toBeNull();
    expect(
      check("weak-password-no-mfa", person({ mfa: false, password: { ...base, days: null } })),
    ).toBeNull();
  });
});

describe("admin-weak-device", () => {
  it("trips for a device that is not compliant or not encrypted, naming it", () => {
    expect(
      check(
        "admin-weak-device",
        person({
          roles: admin(),
          devices: [{ name: "LAP-1", compliance: "Not compliant", encrypted: true }],
        }),
      )?.reason,
    ).toBe("Holds Global Administrator (permanent) and uses LAP-1, which is not compliant.");
    expect(
      check(
        "admin-weak-device",
        person({
          roles: admin(),
          devices: [{ name: "LAP-2", compliance: "Compliant", encrypted: false }],
        }),
      )?.reason,
    ).toContain("LAP-2, which is not encrypted");
    expect(
      check(
        "admin-weak-device",
        person({
          roles: admin(),
          devices: [{ name: null, compliance: "Not compliant", encrypted: false }],
        }),
      )?.reason,
    ).toContain("a managed device, which is not compliant and not encrypted");
  });

  it("stays quiet for a healthy, grace-period or unknown device, or a non-administrator", () => {
    const device = (compliance: string | null, encrypted: boolean | null) => [
      { name: "LAP", compliance, encrypted },
    ];
    expect(
      check("admin-weak-device", person({ roles: admin(), devices: device("Compliant", true) })),
    ).toBeNull();
    expect(
      check(
        "admin-weak-device",
        person({ roles: admin(), devices: device("In grace period", null) }),
      ),
    ).toBeNull();
    expect(
      check("admin-weak-device", person({ roles: admin(), devices: device("Unknown", null) })),
    ).toBeNull();
    expect(
      check(
        "admin-weak-device",
        person({ roles: helpdesk, devices: device("Not compliant", false) }),
      ),
    ).toBeNull();
    expect(check("admin-weak-device", person({ roles: admin() }))).toBeNull();
  });
});

describe("disabled-still-holding", () => {
  it("trips for a disabled account with a role, licenses or owned groups", () => {
    expect(
      check("disabled-still-holding", person({ enabled: false, roles: helpdesk }))?.reason,
    ).toBe("The account is disabled but still holds Helpdesk Administrator (permanent).");
    expect(
      check("disabled-still-holding", person({ enabled: false, licenses: ["E3", "E5"] }))?.reason,
    ).toBe("The account is disabled but still holds 2 licenses.");
    expect(
      check("disabled-still-holding", person({ enabled: false, ownedGroups: ["g"] }))?.reason,
    ).toBe("The account is disabled but still holds ownership of 1 group.");
  });

  it("stays quiet for an enabled or unknown account, and for a disabled one holding nothing", () => {
    expect(
      check("disabled-still-holding", person({ enabled: true, roles: helpdesk, licenses: ["E3"] })),
    ).toBeNull();
    expect(check("disabled-still-holding", person({ enabled: null, roles: helpdesk }))).toBeNull();
    expect(check("disabled-still-holding", person({ enabled: false }))).toBeNull();
  });
});

describe("paying-for-nothing", () => {
  it("trips for a licensed account that is quiet everywhere or never signed in", () => {
    expect(
      check("paying-for-nothing", person({ licenses: ["E3"], verdict: "Quiet everywhere" }))
        ?.reason,
    ).toBe("Holds 1 license and has been quiet everywhere.");
    expect(
      check("paying-for-nothing", person({ licenses: ["E3", "E5"], verdict: "Never signed in" }))
        ?.reason,
    ).toBe("Holds 2 licenses and has never signed in.");
  });

  it("stays quiet for other verdicts, unlicensed accounts, and shared, room or equipment mailboxes", () => {
    for (const verdict of [
      "Active",
      "Recently quiet",
      "Signs in, uses nothing",
      "Not enough data",
      null,
    ])
      expect(check("paying-for-nothing", person({ licenses: ["E3"], verdict }))).toBeNull();
    expect(check("paying-for-nothing", person({ verdict: "Quiet everywhere" }))).toBeNull();
    for (const mailboxType of [
      "SharedMailbox",
      "RoomMailbox",
      "EquipmentMailbox",
      "Shared mailbox",
    ])
      expect(
        check(
          "paying-for-nothing",
          person({ licenses: ["E3"], verdict: "Quiet everywhere", mailboxType }),
        ),
      ).toBeNull();
    expect(
      check(
        "paying-for-nothing",
        person({ licenses: ["E3"], verdict: "Quiet everywhere", mailboxType: "UserMailbox" }),
      ),
    ).not.toBeNull();
  });
});

describe("Conditional Access reach", () => {
  const who = { id: "U-1", groupIds: ["g-staff"], roleIds: ["role-1"], guest: false };

  it("covers a person included by all users, id, direct group or role", () => {
    for (const covering of [
      policy(),
      policy({ includeUsers: ["u-1"] }),
      policy({ includeUsers: [], includeGroups: ["g-staff"] }),
      policy({ includeUsers: [], includeRoles: ["role-1"] }),
    ])
      expect(mfaReach(who, [covering])).toEqual({ state: "covered" });
  });

  it("leaves a person out by id, direct group or role", () => {
    for (const leaving of [
      policy({ excludeUsers: ["u-1"] }),
      policy({ excludeGroups: ["g-staff"] }),
      policy({ excludeRoles: ["role-1"] }),
    ])
      expect(mfaReach(who, [leaving])).toEqual({
        state: "left-out",
        policy: "Require MFA for everyone",
      });
  });

  it("is covered when another policy covers a person one policy leaves out", () => {
    expect(
      mfaReach(who, [
        policy({ excludeUsers: ["u-1"] }),
        policy({ id: "p2", includeUsers: [], includeGroups: ["g-staff"] }),
      ]),
    ).toEqual({ state: "covered" });
  });

  it("is none when no policy is aimed at the person, or there are no policies", () => {
    expect(mfaReach(who, [policy({ includeUsers: [], includeGroups: ["g-other"] })])).toEqual({
      state: "none",
    });
    expect(mfaReach(who, [])).toEqual({ state: "none" });
  });

  it("matches nested groups never: only direct groups count", () => {
    expect(
      mfaReach({ ...who, groupIds: ["g-child"] }, [
        policy({ includeUsers: [], includeGroups: ["g-parent"] }),
      ]),
    ).toEqual({ state: "none" });
  });

  it("matches no group when memberships were not read, in either direction", () => {
    const unknown = { ...who, groupIds: null };
    expect(mfaReach(unknown, [policy({ excludeGroups: ["g-staff"] })])).toEqual({
      state: "covered",
    });
    expect(mfaReach(unknown, [policy({ includeUsers: [], includeGroups: ["g-staff"] })])).toEqual({
      state: "none",
    });
  });

  it("matches guests by their selection", () => {
    const guest = { ...who, guest: true };
    expect(mfaReach(guest, [policy({ includeUsers: [], includesGuests: true })])).toEqual({
      state: "covered",
    });
    expect(mfaReach(guest, [policy({ excludesGuests: true })])).toEqual({
      state: "left-out",
      policy: "Require MFA for everyone",
    });
    expect(mfaReach(who, [policy({ includeUsers: [], includesGuests: true })])).toEqual({
      state: "none",
    });
  });

  it("names the first left-out policy by name", () => {
    expect(
      mfaReach(who, [
        policy({ name: "B", excludeUsers: ["u-1"] }),
        policy({ id: "p2", name: "A", excludeUsers: ["u-1"] }),
      ]),
    ).toEqual({ state: "left-out", policy: "A" });
  });

  it("reads only enabled policies that require MFA", () => {
    const raw = (state: string, grant: unknown) => ({
      id: state,
      displayName: `Policy ${state}`,
      state,
      conditions: { users: { includeUsers: ["All"], excludeGroups: ["G-1"] } },
      grantControls: grant,
    });
    const read = mfaPolicies([
      raw("enabled", { builtInControls: ["mfa"] }),
      raw("enabledForReportingButNotEnforced", { builtInControls: ["mfa"] }),
      raw("disabled", { builtInControls: ["mfa"] }),
      raw("enabled", { builtInControls: ["block"] }),
      {
        ...raw("enabled", null),
        id: "strength",
        grantControls: { builtInControls: [], authenticationStrength: { displayName: "x" } },
      },
      { id: 5 },
    ]);
    expect(read.policies.map((item) => item.id)).toEqual(["enabled", "strength"]);
    expect(read.skipped).toBe(1);
    expect(read.policies[0]?.policy.conditions.users).toMatchObject({
      includeUsers: ["All"],
      excludeGroups: ["G-1"],
    });
  });
});
