import { describe, expect, it } from "vitest";
import { mfaPolicies, mfaReach, policyReach, policySchema } from "../src/index.ts";

// Whether a Conditional Access policy reaches one person, from its "Users" part alone. Synthetic
// ids only.
const person = {
  id: "aaaaaaaa-0000-4000-8000-000000000001",
  userType: "Member",
  groupIds: ["bbbbbbbb-0000-4000-8000-000000000001", "bbbbbbbb-0000-4000-8000-000000000002"],
  roleTemplateIds: ["cccccccc-0000-4000-8000-000000000001"],
};
const other = "aaaaaaaa-0000-4000-8000-0000000000ff";

function policy(users: Record<string, unknown>, state = "enabled") {
  return policySchema.parse({
    id: "policy-1",
    displayName: "A policy",
    state,
    conditions: { users },
    grantControls: { operator: "OR", builtInControls: ["mfa"] },
  });
}

describe("policyReach", () => {
  it("reaches everyone when the policy includes all users", () => {
    const reach = policyReach(policy({ includeUsers: ["All"] }), person);
    expect(reach).toEqual({
      effect: "applies",
      targeted: true,
      included: [{ kind: "all", id: null }],
      excluded: [],
    });
  });

  it("reaches the person named by id", () => {
    const reach = policyReach(policy({ includeUsers: [other, person.id] }), person);
    expect(reach.effect).toBe("applies");
    expect(reach.included).toEqual([{ kind: "person", id: person.id }]);
  });

  it("reaches a person through a group they are in, direct or nested", () => {
    // The caller passes every group the person is in, so a nested group is just another id.
    const direct = policyReach(
      policy({ includeUsers: [], includeGroups: ["bbbbbbbb-0000-4000-8000-000000000001"] }),
      person,
    );
    const nested = policyReach(
      policy({ includeUsers: [], includeGroups: ["bbbbbbbb-0000-4000-8000-000000000002"] }),
      person,
    );
    expect(direct.included).toEqual([
      { kind: "group", id: "bbbbbbbb-0000-4000-8000-000000000001" },
    ]);
    expect(nested.included).toEqual([
      { kind: "group", id: "bbbbbbbb-0000-4000-8000-000000000002" },
    ]);
  });

  it("reaches a person through a role they hold", () => {
    const reach = policyReach(
      policy({ includeUsers: [], includeRoles: ["CCCCCCCC-0000-4000-8000-000000000001"] }),
      person,
    );
    expect(reach.included).toEqual([{ kind: "role", id: "CCCCCCCC-0000-4000-8000-000000000001" }]);
  });

  it("reaches guests only when the person is a guest", () => {
    const rule = {
      includeUsers: [],
      includeGuestsOrExternalUsers: { guestOrExternalUserTypes: "b2bCollaborationGuest" },
    };
    expect(policyReach(policy(rule), person).effect).toBe("none");
    expect(policyReach(policy(rule), { ...person, userType: "Guest" }).included).toEqual([
      { kind: "guests", id: null },
    ]);
    const legacy = policyReach(policy({ includeUsers: ["GuestsOrExternalUsers"] }), {
      ...person,
      userType: "Guest",
    });
    expect(legacy.effect).toBe("applies");
    expect(policyReach(policy({ includeUsers: ["GuestsOrExternalUsers"] }), person).effect).toBe(
      "none",
    );
  });

  it("lets an exclusion beat an inclusion", () => {
    const reach = policyReach(
      policy({
        includeUsers: ["All"],
        excludeGroups: ["bbbbbbbb-0000-4000-8000-000000000002"],
      }),
      person,
    );
    expect(reach.effect).toBe("excluded");
    expect(reach.targeted).toBe(true);
    expect(reach.excluded).toEqual([{ kind: "group", id: "bbbbbbbb-0000-4000-8000-000000000002" }]);
  });

  it("reports an exclusion by person, role or guests", () => {
    expect(
      policyReach(policy({ includeUsers: ["All"], excludeUsers: [person.id] }), person).excluded,
    ).toEqual([{ kind: "person", id: person.id }]);
    expect(
      policyReach(
        policy({ includeUsers: ["All"], excludeRoles: ["cccccccc-0000-4000-8000-000000000001"] }),
        person,
      ).effect,
    ).toBe("excluded");
    expect(
      policyReach(
        policy({
          includeUsers: ["All"],
          excludeGuestsOrExternalUsers: { guestOrExternalUserTypes: "b2bCollaborationGuest" },
        }),
        { ...person, userType: "Guest" },
      ).effect,
    ).toBe("excluded");
  });

  it("names an exclusion even when no inclusion reached the person", () => {
    const reach = policyReach(policy({ includeUsers: [other], excludeUsers: [person.id] }), person);
    expect(reach).toMatchObject({ effect: "excluded", targeted: false });
  });

  it("matches report-only and off policies the same as enabled ones", () => {
    // The state is the caller's to show; reach does not depend on it.
    for (const state of ["enabledForReportingButNotEnforced", "disabled"])
      expect(policyReach(policy({ includeUsers: ["All"] }, state), person).effect).toBe("applies");
  });

  it("does not reach a person a policy names nobody they match", () => {
    const reach = policyReach(
      policy({
        includeUsers: [other],
        includeGroups: ["bbbbbbbb-0000-4000-8000-0000000000ff"],
        includeRoles: ["cccccccc-0000-4000-8000-0000000000ff"],
      }),
      person,
    );
    expect(reach).toEqual({ effect: "none", targeted: false, included: [], excluded: [] });
  });

  it("reads None as nobody and an unknown user type as not a guest", () => {
    expect(policyReach(policy({ includeUsers: ["None"] }), person).effect).toBe("none");
    expect(
      policyReach(policy({ includeUsers: ["GuestsOrExternalUsers"] }), {
        ...person,
        userType: null,
      }).effect,
    ).toBe("none");
  });
});

// The Combined risks audit decides reach with policyReach too, so the two cannot drift apart.
describe("policyReach and the Combined risks check agree", () => {
  const role = "cccccccc-0000-4000-8000-000000000001";
  const group = "bbbbbbbb-0000-4000-8000-000000000001";
  const subject = { id: person.id, groupIds: [group], roleIds: [role], guest: false };
  const cases: [string, Record<string, unknown>, "applies" | "excluded"][] = [
    ["an include by group", { includeUsers: [], includeGroups: [group] }, "applies"],
    ["an exclude by group", { includeUsers: ["All"], excludeGroups: [group] }, "excluded"],
    ["an exclude by role", { includeUsers: ["All"], excludeRoles: [role] }, "excluded"],
  ];

  it.each(cases)("%s", (_name, users, effect) => {
    const raw = {
      id: "policy-1",
      displayName: "A policy",
      state: "enabled",
      conditions: { users },
      grantControls: { operator: "OR", builtInControls: ["mfa"] },
    };
    const reach = policyReach(policySchema.parse(raw), {
      ...person,
      groupIds: subject.groupIds,
      roleTemplateIds: subject.roleIds,
    });
    const { policies } = mfaPolicies([raw]);
    expect(reach.effect).toBe(effect);
    expect(mfaReach(subject, policies)).toEqual(
      effect === "applies" ? { state: "covered" } : { state: "left-out", policy: "A policy" },
    );
  });
});
