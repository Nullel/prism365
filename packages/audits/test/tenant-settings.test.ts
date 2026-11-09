import type { AuditRow } from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import {
  type AuditContext,
  type AuditResult,
  auditDefinition,
  builtInAudits,
  createConditionalAccessAudit,
  REVIEW_ORDER,
  SECURITY_DEFAULTS_KEY,
  SECURITY_DEFAULTS_ON,
  TENANT_SETTING_AREAS,
  tenantSettingsAudit,
  type StoredRun,
} from "../src/index.ts";
import { daysAgo, expectValidTable, fake, find, type Route } from "./apps-devices-usage-support.ts";

// Synthetic tenant policies: the shapes Microsoft documents, with invented policy ids and names.

const DEFAULTS = "/v1.0/policies/identitySecurityDefaultsEnforcementPolicy";
const METHODS = "/v1.0/policies/authenticationMethodsPolicy";
const AUTHORIZATION = "/v1.0/policies/authorizationPolicy";
const WORKFLOW = "/v1.0/policies/adminConsentRequestPolicy";
const GRANTS = "/v1.0/policies/permissionGrantPolicies";
const PATHS = [DEFAULTS, METHODS, AUTHORIZATION, WORKFLOW, GRANTS];

const RESTRICTED = "2af84b1e-32c8-42b7-82bc-daa82404023b";
const LIMITED = "10dae51f-b6af-4016-8d66-8c2a99b929b3";
const SAME_AS_MEMBERS = "a0b1b346-4d3e-4e8b-98f8-753987be4970";
const LOW = "ManagePermissionGrantsForSelf.microsoft-user-default-low";
const LEGACY = "ManagePermissionGrantsForSelf.microsoft-user-default-legacy";

type Json = Record<string, unknown>;

interface Bodies {
  defaults: Json;
  methods: Json;
  authorization: Json;
  workflow: Json;
  grants: Json[];
}

const method = (id: string, state: string) => ({
  "@odata.type": `#microsoft.graph.${id}AuthenticationMethodConfiguration`,
  id,
  state,
});

// A tenant that meets every expectation.
function healthy(): Bodies {
  return {
    defaults: { id: "00000000-0000-0000-0000-000000000005", isEnabled: true },
    methods: {
      policyMigrationState: "migrationComplete",
      registrationEnforcement: {
        authenticationMethodsRegistrationCampaign: { state: "enabled" },
      },
      authenticationMethodConfigurations: [
        method("Sms", "disabled"),
        method("Voice", "disabled"),
        method("Email", "enabled"),
        method("Fido2", "enabled"),
        method("MicrosoftAuthenticator", "enabled"),
        method("TemporaryAccessPass", "enabled"),
        method("SoftwareOath", "enabled"),
        method("X509Certificate", "disabled"),
      ],
    },
    authorization: {
      allowInvitesFrom: "adminsAndGuestInviters",
      guestUserRoleId: RESTRICTED,
      allowEmailVerifiedUsersToJoinOrganization: false,
      allowedToSignUpEmailBasedSubscriptions: false,
      allowedToUseSSPR: true,
      blockMsolPowerShell: true,
      defaultUserRolePermissions: {
        allowedToCreateApps: false,
        allowedToCreateSecurityGroups: true,
        allowedToReadOtherUsers: true,
        permissionGrantPoliciesAssigned: [LOW],
      },
    },
    workflow: { isEnabled: true, version: 1 },
    grants: [
      { id: "microsoft-user-default-low", displayName: "Low impact from verified publishers" },
      { id: "microsoft-user-default-legacy", displayName: "Legacy user consent" },
    ],
  };
}

function withMethod(body: Bodies, id: string, state: string | null): Bodies {
  const list = (body.methods.authenticationMethodConfigurations as Json[]).filter(
    (item) => item.id !== id,
  );
  if (state !== null) list.push(method(id, state));
  return { ...body, methods: { ...body.methods, authenticationMethodConfigurations: list } };
}

function withAuthorization(body: Bodies, over: Json, role: Json = {}): Bodies {
  const base = body.authorization;
  return {
    ...body,
    authorization: {
      ...base,
      ...over,
      defaultUserRolePermissions: {
        ...(base.defaultUserRolePermissions as Json),
        ...role,
      },
    },
  };
}

function routeOf(bodies: Bodies, failing: Record<string, number> = {}): Route {
  return (url) => {
    const status = failing[url.pathname];
    if (status) return new Response("{}", { status });
    switch (url.pathname) {
      case DEFAULTS:
        return Response.json(bodies.defaults);
      case METHODS:
        return Response.json(bodies.methods);
      case AUTHORIZATION:
        return Response.json(bodies.authorization);
      case WORKFLOW:
        return Response.json(bodies.workflow);
      case GRANTS:
        return Response.json({ value: bodies.grants });
      default:
        return undefined;
    }
  };
}

function storedWith(run: Partial<StoredRun> | null) {
  const asked: string[] = [];
  return {
    asked,
    async latestRun(auditId: string) {
      asked.push(auditId);
      return run ? { finishedAt: daysAgo(1), complete: true, scope: null, rows: [], ...run } : null;
    },
    async openFindings() {
      return [];
    },
  };
}

interface Setup {
  failing?: Record<string, number>;
  run?: Partial<StoredRun> | null;
  noStored?: boolean;
}

async function run(bodies: Bodies, setup: Setup = {}) {
  const network = fake(tenantSettingsAudit, routeOf(bodies, setup.failing));
  const stored = storedWith(setup.run ?? null);
  const context: AuditContext = setup.noStored ? network.context : { ...network.context, stored };
  const result = await tenantSettingsAudit.run(context);
  return { result, requests: network.requests, stored };
}

const cell = (result: AuditResult, key: string, name: string) => find(result.rows, key).cells[name];
const status = (result: AuditResult, key: string) => cell(result, key, "status");
const value = (result: AuditResult, key: string) => String(cell(result, key, "value"));

// The rows a Conditional Access run holds, as that audit writes them.
const caRow = (over: { tags?: string[]; appliesTo?: string; policy?: string } = {}): AuditRow => ({
  key: "ca-1",
  personId: null,
  tags: over.tags ?? ["enabled", "mfa"],
  cells: {
    policy: over.policy ?? "Require MFA for everyone",
    appliesTo: over.appliesTo ?? "All users",
  },
});

const off = (bodies: Bodies): Bodies => ({ ...bodies, defaults: { isEnabled: false } });

describe("the published definition", () => {
  const definition = auditDefinition(tenantSettingsAudit);

  it("reads the five policies under the tenant policies pack, with stored results, unscoped", () => {
    expect(tenantSettingsAudit.packs).toEqual(["identity.policies.read"]);
    expect(tenantSettingsAudit.optionalPacks).toEqual([]);
    expect(definition).toMatchObject({
      id: "tenant.settings",
      permissions: ["Policy.Read.All"],
      optionalPermissions: [],
      scoped: false,
      handoff: false,
    });
    expect(tenantSettingsAudit.usesStoredResults).toBe(true);
    expect(builtInAudits).toContain(tenantSettingsAudit);
  });

  it("has the views, columns and tiles the contract names", () => {
    expect(definition.views.map((view) => view.id)).toEqual([
      "all",
      "review",
      "meets",
      "not-read",
      "security-defaults",
      "authentication-methods",
      "guests",
      "applications",
      "consent",
      "users",
    ]);
    expect(definition.columns.map((column) => column.key)).toEqual([
      "setting",
      "area",
      "value",
      "expected",
      "status",
      "why",
      "path",
    ]);
    const facets = definition.columns.filter((column) => column.facet).map((column) => column.key);
    expect(facets).toEqual(["area", "status"]);
    expect(definition.columns.find((column) => column.key === "path")?.hidden).toBe(true);
    expect(definition.defaultSort).toEqual({ column: "status", direction: "desc" });
  });

  it("gives a valid table over a healthy tenant", async () => {
    const result = await expectValidTable(tenantSettingsAudit, routeOf(healthy()), {
      granted: ["Policy.Read.All"],
    });
    expect(result.summary.map((tile) => [tile.label, tile.tone, tile.view])).toEqual([
      ["Review", "warning", "review"],
      ["Meets", "good", "meets"],
      ["Not read", "neutral", "not-read"],
    ]);
  });

  it("sits in the summary stage, after Conditional Access and before Combined risks", () => {
    const ids = REVIEW_ORDER.map((step) => step.auditId);
    expect(REVIEW_ORDER.find((step) => step.auditId === "tenant.settings")?.stage).toBe("summary");
    expect(ids.indexOf("tenant.settings")).toBeGreaterThan(
      ids.indexOf("identity.conditional-access"),
    );
    expect(ids.indexOf("tenant.settings")).toBeLessThan(ids.indexOf("identity.combined-risks"));
  });
});

describe("what it asks Microsoft", () => {
  it("makes one request per path, five in all, and nothing else", async () => {
    const { requests } = await run(healthy());
    expect(requests.map((url) => url.pathname).sort()).toEqual([...PATHS].sort());
  });

  it("reads the stored Conditional Access run only when security defaults are off", async () => {
    const on = await run(healthy());
    expect(on.stored.asked).toEqual([]);
    const disabled = await run(off(healthy()));
    expect(disabled.stored.asked).toEqual(["identity.conditional-access"]);
  });
});

describe("a healthy tenant", () => {
  it("has every row Meets, none Review or Not read, and a complete run", async () => {
    const { result } = await run(healthy());
    expect(result.rows).toHaveLength(22);
    for (const item of result.rows) expect(item.cells.status, item.key).toBe("Meets");
    expect(result.summary.map((tile) => tile.value)).toEqual([0, 22, 0]);
    expect(result.complete).toBe(true);
  });

  it("covers each area and gives every row an area tag, a status tag and one setting", async () => {
    const { result } = await run(healthy());
    const areas = new Set(result.rows.map((item) => item.cells.area));
    expect(areas).toEqual(new Set(TENANT_SETTING_AREAS));
    for (const item of result.rows) {
      expect(item.tags).toContain("meets");
      expect(item.personId).toBeNull();
    }
    expect(new Set(result.rows.map((item) => item.key)).size).toBe(result.rows.length);
  });
});

// Each row, both ways: the value that meets and the value that needs review.
interface Case {
  key: string;
  area: string;
  meets: (body: Bodies) => Bodies;
  review: (body: Bodies) => Bodies;
  shows: string;
}

const cases: Case[] = [
  {
    key: "method-sms",
    area: "Authentication methods",
    meets: (body) => withMethod(body, "Sms", "disabled"),
    review: (body) => withMethod(body, "Sms", "enabled"),
    shows: "Enabled",
  },
  {
    key: "method-voice",
    area: "Authentication methods",
    meets: (body) => withMethod(body, "Voice", "disabled"),
    review: (body) => withMethod(body, "Voice", "enabled"),
    shows: "Enabled",
  },
  {
    key: "method-fido2",
    area: "Authentication methods",
    meets: (body) => withMethod(body, "Fido2", "enabled"),
    review: (body) => withMethod(body, "Fido2", "disabled"),
    shows: "Disabled",
  },
  {
    key: "method-authenticator",
    area: "Authentication methods",
    meets: (body) => withMethod(body, "MicrosoftAuthenticator", "enabled"),
    review: (body) => withMethod(body, "MicrosoftAuthenticator", "disabled"),
    shows: "Disabled",
  },
  {
    key: "method-tap",
    area: "Authentication methods",
    meets: (body) => withMethod(body, "TemporaryAccessPass", "enabled"),
    review: (body) => withMethod(body, "TemporaryAccessPass", "disabled"),
    shows: "Disabled",
  },
  {
    key: "registration-campaign",
    area: "Authentication methods",
    meets: (body) => body,
    review: (body) => ({
      ...body,
      methods: {
        ...body.methods,
        registrationEnforcement: {
          authenticationMethodsRegistrationCampaign: { state: "disabled" },
        },
      },
    }),
    shows: "Disabled",
  },
  {
    key: "policy-migration",
    area: "Authentication methods",
    meets: (body) => body,
    review: (body) => ({
      ...body,
      methods: { ...body.methods, policyMigrationState: "migrationInProgress" },
    }),
    shows: "In progress",
  },
  {
    key: "guest-invites",
    area: "Guests",
    meets: (body) => withAuthorization(body, { allowInvitesFrom: "none" }),
    review: (body) => withAuthorization(body, { allowInvitesFrom: "everyone" }),
    shows: "Everyone, including guests",
  },
  {
    key: "guest-role",
    area: "Guests",
    meets: (body) => withAuthorization(body, { guestUserRoleId: RESTRICTED }),
    review: (body) => withAuthorization(body, { guestUserRoleId: LIMITED }),
    shows: "Limited access",
  },
  {
    key: "email-verified-join",
    area: "Users",
    meets: (body) => withAuthorization(body, { allowEmailVerifiedUsersToJoinOrganization: false }),
    review: (body) => withAuthorization(body, { allowEmailVerifiedUsersToJoinOrganization: true }),
    shows: "Allowed",
  },
  {
    key: "email-subscriptions",
    area: "Users",
    meets: (body) => withAuthorization(body, { allowedToSignUpEmailBasedSubscriptions: false }),
    review: (body) => withAuthorization(body, { allowedToSignUpEmailBasedSubscriptions: true }),
    shows: "Allowed",
  },
  {
    key: "create-apps",
    area: "Applications",
    meets: (body) => withAuthorization(body, {}, { allowedToCreateApps: false }),
    review: (body) => withAuthorization(body, {}, { allowedToCreateApps: true }),
    shows: "Allowed",
  },
  {
    key: "user-consent",
    area: "Consent",
    meets: (body) => withAuthorization(body, {}, { permissionGrantPoliciesAssigned: [] }),
    review: (body) => withAuthorization(body, {}, { permissionGrantPoliciesAssigned: [LEGACY] }),
    shows: "On: Legacy user consent",
  },
  {
    key: "admin-consent-workflow",
    area: "Consent",
    meets: (body) => ({ ...body, workflow: { isEnabled: true } }),
    review: (body) => ({ ...body, workflow: { isEnabled: false } }),
    shows: "Off",
  },
];

describe("each setting against its expected value", () => {
  for (const item of cases) {
    it(`${item.key} meets when it is as expected`, async () => {
      const { result } = await run(item.meets(healthy()));
      expect(status(result, item.key)).toBe("Meets");
      expect(find(result.rows, item.key).tags).toEqual(
        expect.arrayContaining(["meets", item.area.toLowerCase().replaceAll(" ", "-")]),
      );
    });

    it(`${item.key} needs review when it differs`, async () => {
      const { result } = await run(item.review(healthy()));
      expect(status(result, item.key)).toBe("Review");
      expect(value(result, item.key)).toBe(item.shows);
      expect(cell(result, item.key, "area")).toBe(item.area);
      expect(find(result.rows, item.key).tags).toContain("review");
      expect(result.summary[0]?.value).toBe(1);
    });
  }

  it("never calls a difference a failure", async () => {
    const worst = withAuthorization(
      withMethod(healthy(), "Sms", "enabled"),
      { allowInvitesFrom: "everyone" },
      { allowedToCreateApps: true },
    );
    const { result } = await run(worst);
    const words = JSON.stringify([result.rows, result.notes, result.summary]);
    expect(words).not.toMatch(/\bfail/i);
  });
});

describe("settings that are only reported", () => {
  const reportOnly = [
    "method-email",
    "method-oath",
    "method-x509",
    "create-security-groups",
    "read-other-users",
    "self-service-password-reset",
    "msol-powershell",
  ];

  it("meet with either value, and show the value Microsoft gave", async () => {
    const flipped = withMethod(
      withMethod(withMethod(healthy(), "Email", "disabled"), "SoftwareOath", "disabled"),
      "X509Certificate",
      "enabled",
    );
    const body = withAuthorization(
      flipped,
      { allowedToUseSSPR: false, blockMsolPowerShell: false },
      { allowedToCreateSecurityGroups: false, allowedToReadOtherUsers: false },
    );
    const { result } = await run(body);
    for (const key of reportOnly) {
      expect(status(result, key), key).toBe("Meets");
      expect(cell(result, key, "expected"), key).toBe("Your choice");
    }
    expect(value(result, "method-email")).toBe("Disabled");
    expect(value(result, "method-x509")).toBe("Enabled");
    expect(value(result, "self-service-password-reset")).toBe("Not allowed");
    expect(value(result, "msol-powershell")).toBe("Not blocked");
    expect(value(result, "create-security-groups")).toBe("Not allowed");
  });

  it("are Not read, never Meets, when Microsoft left them out", async () => {
    const body = healthy();
    delete body.authorization.allowedToUseSSPR;
    delete body.authorization.blockMsolPowerShell;
    (body.authorization.defaultUserRolePermissions as Json).allowedToCreateSecurityGroups =
      undefined;
    const bare = withMethod(withMethod(body, "Email", null), "X509Certificate", null);
    const { result } = await run(bare);
    for (const key of [
      "self-service-password-reset",
      "msol-powershell",
      "create-security-groups",
      "method-email",
      "method-x509",
    ])
      expect(status(result, key), key).toBe("Not read");
  });
});

describe("the guest role", () => {
  it("names the three roles Microsoft documents by id", async () => {
    for (const [id, words] of [
      [RESTRICTED, "Restricted access"],
      [LIMITED, "Limited access"],
      [SAME_AS_MEMBERS, "Same access as members"],
    ] as const) {
      const { result } = await run(withAuthorization(healthy(), { guestUserRoleId: id }));
      expect(value(result, "guest-role")).toBe(words);
      expect(status(result, "guest-role")).toBe(id === RESTRICTED ? "Meets" : "Review");
    }
  });

  it("does not guess at a role it does not know, and matches ids in any case", async () => {
    const unknown = await run(
      withAuthorization(healthy(), { guestUserRoleId: "00000000-0000-4000-8000-000000000000" }),
    );
    expect(status(unknown.result, "guest-role")).toBe("Review");
    expect(value(unknown.result, "guest-role")).toBe("A role Prism365 does not recognise");
    const upper = await run(
      withAuthorization(healthy(), { guestUserRoleId: RESTRICTED.toUpperCase() }),
    );
    expect(status(upper.result, "guest-role")).toBe("Meets");
  });

  it("says Review for who can invite guests when members or everyone can", async () => {
    for (const [from, words, state] of [
      ["none", "No one", "Meets"],
      ["adminsAndGuestInviters", "Only administrators and guest inviters", "Meets"],
      [
        "adminsGuestInvitersAndAllMembers",
        "Administrators, guest inviters and all members",
        "Review",
      ],
      ["everyone", "Everyone, including guests", "Review"],
    ] as const) {
      const { result } = await run(withAuthorization(healthy(), { allowInvitesFrom: from }));
      expect(value(result, "guest-invites")).toBe(words);
      expect(status(result, "guest-invites")).toBe(state);
    }
  });
});

describe("authentication methods", () => {
  it("reads each method by its id from the policy, ignoring case", async () => {
    const body = healthy();
    body.methods.authenticationMethodConfigurations = [
      { id: "sms", state: "Enabled" },
      { id: "FIDO2", state: "DISABLED" },
    ];
    const { result } = await run(body);
    expect(value(result, "method-sms")).toBe("Enabled");
    expect(status(result, "method-sms")).toBe("Review");
    expect(value(result, "method-fido2")).toBe("Disabled");
    expect(status(result, "method-voice")).toBe("Not read");
  });

  it("is Not read for a method Microsoft did not list or gave an unknown state", async () => {
    const body = withMethod(withMethod(healthy(), "Sms", null), "Voice", "somethingNew");
    const { result } = await run(body);
    expect(status(result, "method-sms")).toBe("Not read");
    expect(value(result, "method-sms")).toBe("Microsoft did not return this setting");
    expect(status(result, "method-voice")).toBe("Not read");
  });

  it("treats a registration campaign Microsoft manages as Review, and reads other migration states", async () => {
    const managed = healthy();
    managed.methods.registrationEnforcement = {
      authenticationMethodsRegistrationCampaign: { state: "default" },
    };
    managed.methods.policyMigrationState = "preMigration";
    const { result } = await run(managed);
    expect(status(result, "registration-campaign")).toBe("Review");
    expect(value(result, "registration-campaign")).toBe("Managed by Microsoft");
    expect(status(result, "policy-migration")).toBe("Review");
    expect(value(result, "policy-migration")).toBe("Not started");
  });

  it("is Not read when the campaign or the migration state is missing", async () => {
    const body = healthy();
    delete body.methods.registrationEnforcement;
    delete body.methods.policyMigrationState;
    const { result } = await run(body);
    expect(status(result, "registration-campaign")).toBe("Not read");
    expect(status(result, "policy-migration")).toBe("Not read");
  });
});

describe("security defaults", () => {
  it("meets when they are on, says On in the way Combined risks reads, and needs no other audit", async () => {
    const { result, stored } = await run(healthy());
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Meets");
    expect(value(result, SECURITY_DEFAULTS_KEY)).toBe(SECURITY_DEFAULTS_ON);
    expect(stored.asked).toEqual([]);
  });

  it("meets when off with a fresh Conditional Access run that has an enabled MFA policy for all users", async () => {
    const { result } = await run(off(healthy()), {
      run: { finishedAt: daysAgo(2), rows: [caRow()] },
    });
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Meets");
    expect(value(result, SECURITY_DEFAULTS_KEY)).toBe(
      "Off, with a Conditional Access policy that requires MFA for all users",
    );
    // The tag marks security defaults themselves; a policy standing in for them is not the same.
    expect(value(result, SECURITY_DEFAULTS_KEY)).not.toBe(SECURITY_DEFAULTS_ON);
    const note = result.notes.find((item) => item.startsWith("Security defaults are off"));
    expect(note).toContain("run of 2026-10-20");
    expect(note).toContain("Require MFA for everyone");
  });

  it("names the exclusions of the policy it relied on", async () => {
    const { result } = await run(off(healthy()), {
      run: { rows: [caRow({ tags: ["enabled", "mfa", "exclusions"] })] },
    });
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Meets");
    expect(result.notes.join("\n")).toContain("lists exclusions");
    expect(result.notes.join("\n")).toContain(
      "Apps, conditions and grants that accept another control in place of MFA are not weighed either.",
    );
  });

  it("needs review when the fresh run has no such policy", async () => {
    const { result } = await run(off(healthy()), {
      run: {
        rows: [
          caRow({ tags: ["report-only", "mfa"] }),
          caRow({ tags: ["enabled"] }),
          caRow({ appliesTo: "2 groups" }),
          caRow({ tags: ["mfa"] }),
        ],
      },
    });
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Review");
    expect(value(result, SECURITY_DEFAULTS_KEY)).toBe(
      "Off. No enabled Conditional Access policy requires MFA for all users",
    );
    expect(result.notes.join("\n")).toContain("run of 2026-10-21");
  });

  it("says when the run it relied on was partial", async () => {
    const { result } = await run(off(healthy()), { run: { complete: false, rows: [] } });
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Review");
    expect(result.notes.join("\n")).toContain("partial result");
  });

  it("asks for the Conditional Access run when there is none", async () => {
    const never = await run(off(healthy()), { run: null });
    expect(status(never.result, SECURITY_DEFAULTS_KEY)).toBe("Review");
    expect(value(never.result, SECURITY_DEFAULTS_KEY)).toBe(
      "Off. Run Conditional Access policies first",
    );
    expect(never.result.notes.join("\n")).toContain("has never run");
    const without = await run(off(healthy()), { noStored: true });
    expect(value(without.result, SECURITY_DEFAULTS_KEY)).toBe(
      "Off. Run Conditional Access policies first",
    );
  });

  it("does not use a stale run, and names its date", async () => {
    const { result } = await run(off(healthy()), {
      run: { finishedAt: daysAgo(8), rows: [caRow()] },
    });
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Review");
    expect(value(result, SECURITY_DEFAULTS_KEY)).toBe("Off. Run Conditional Access policies first");
    const note = result.notes.join("\n");
    expect(note).toContain("last ran 2026-10-14");
    expect(note).toContain("more than 7 days ago");
  });

  it("still uses a run that finished exactly seven days ago", async () => {
    const { result } = await run(off(healthy()), {
      run: { finishedAt: daysAgo(7), rows: [caRow()] },
    });
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Meets");
  });

  it("is Not read when Microsoft leaves out isEnabled", async () => {
    const { result } = await run({ ...healthy(), defaults: {} });
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Not read");
  });

  it("matches the words the Conditional Access audit really writes", async () => {
    const policies = [
      {
        id: "policy-1",
        displayName: "Require MFA for everyone",
        state: "enabled",
        conditions: {
          users: { includeUsers: ["All"], excludeUsers: [] },
          applications: { includeApplications: ["All"] },
          clientAppTypes: ["all"],
        },
        grantControls: { operator: "OR", builtInControls: ["mfa"] },
      },
      {
        id: "policy-2",
        displayName: "Require MFA for finance",
        state: "enabled",
        conditions: {
          users: { includeUsers: [], includeGroups: ["g-1"], excludeUsers: [] },
          applications: { includeApplications: ["All"] },
          clientAppTypes: ["all"],
        },
        grantControls: { operator: "OR", builtInControls: ["mfa"] },
      },
    ];
    const audit = createConditionalAccessAudit({ pause: async () => {} });
    const caRun = await audit.run(
      fake(audit, (url) =>
        url.pathname === "/v1.0/identity/conditionalAccess/policies"
          ? Response.json({ value: policies })
          : undefined,
      ).context,
    );
    const everyone = caRun.rows.filter((item) => item.cells.appliesTo === "All users");
    expect(everyone.length).toBe(1);
    const { result } = await run(off(healthy()), { run: { rows: caRun.rows } });
    expect(status(result, SECURITY_DEFAULTS_KEY)).toBe("Meets");

    const onlyGroup = caRun.rows.filter((item) => item.cells.appliesTo !== "All users");
    const narrower = await run(off(healthy()), { run: { rows: onlyGroup } });
    expect(status(narrower.result, SECURITY_DEFAULTS_KEY)).toBe("Review");
  });
});

describe("user consent and the admin consent workflow", () => {
  const consent = (assigned: string[]) =>
    withAuthorization(healthy(), {}, { permissionGrantPoliciesAssigned: assigned });

  it("is off when empty, limited with only the low-impact policy, and open otherwise", async () => {
    const off = await run(consent([]));
    expect(value(off.result, "user-consent")).toBe("Off. Users cannot consent to apps");
    expect(status(off.result, "user-consent")).toBe("Meets");
    const limited = await run(consent([LOW]));
    expect(status(limited.result, "user-consent")).toBe("Meets");
    const both = await run(consent([LOW, LEGACY]));
    expect(status(both.result, "user-consent")).toBe("Review");
    const other = await run(
      consent(["ManagePermissionGrantsForSelf.microsoft-user-default-recommended"]),
    );
    expect(status(other.result, "user-consent")).toBe("Review");
    const lower = await run(consent([LOW.toLowerCase()]));
    expect(status(lower.result, "user-consent")).toBe("Meets");
  });

  it("names the policies from the permission grant policy list", async () => {
    const { result } = await run(consent([LOW, LEGACY]));
    expect(value(result, "user-consent")).toBe(
      "On: Low impact from verified publishers, Legacy user consent",
    );
  });

  it("falls back to what the two long-standing policies mean, then to the id", async () => {
    const body = {
      ...consent([LOW, LEGACY, "ManagePermissionGrantsForSelf.custom-policy"]),
      grants: [],
    };
    const { result } = await run(body);
    expect(value(result, "user-consent")).toBe(
      "On: Verified publishers, low impact, Any app, any permission a user can consent to (legacy), ManagePermissionGrantsForSelf.custom-policy",
    );
  });

  it("keeps the value and marks the run partial when the policy list cannot be read", async () => {
    const { result } = await run(consent([LEGACY]), { failing: { [GRANTS]: 403 } });
    expect(status(result, "user-consent")).toBe("Review");
    expect(value(result, "user-consent")).toContain("(legacy)");
    expect(result.complete).toBe(false);
    expect(result.notes.join("\n")).toContain("permission grant policy list could not be read");
  });

  it("is Not read when the field is missing, and then the workflow is judged without it", async () => {
    const body = healthy();
    (body.authorization.defaultUserRolePermissions as Json).permissionGrantPoliciesAssigned =
      undefined;
    const { result } = await run({ ...body, workflow: { isEnabled: false } });
    expect(status(result, "user-consent")).toBe("Not read");
    expect(status(result, "admin-consent-workflow")).toBe("Review");
    expect(cell(result, "admin-consent-workflow", "expected")).toBe(
      "On, unless users can consent to any app",
    );
  });

  it("expects the workflow when consent is off or limited, and makes it a choice when open", async () => {
    for (const assigned of [[], [LOW]]) {
      const { result } = await run({ ...consent(assigned), workflow: { isEnabled: false } });
      expect(status(result, "admin-consent-workflow")).toBe("Review");
      expect(cell(result, "admin-consent-workflow", "expected")).toBe("On");
      expect(cell(result, "admin-consent-workflow", "why")).toBe(
        "Without the workflow, users have no way to ask for an app to be approved.",
      );
    }
    const open = await run({ ...consent([LEGACY]), workflow: { isEnabled: false } });
    expect(status(open.result, "admin-consent-workflow")).toBe("Meets");
    expect(cell(open.result, "admin-consent-workflow", "expected")).toBe("Your choice");
    const missing = await run({ ...consent([]), workflow: {} });
    expect(status(missing.result, "admin-consent-workflow")).toBe("Not read");
  });
});

describe("owned-resource consent entries", () => {
  const OWNED_TEAMS =
    "ManagePermissionGrantsForOwnedResource.microsoft-dynamically-managed-permissions-for-team";
  const OWNED_CHAT =
    "ManagePermissionGrantsForOwnedResource.microsoft-dynamically-managed-permissions-for-chat";
  const consent = (assigned: unknown) =>
    withAuthorization(healthy(), {}, { permissionGrantPoliciesAssigned: assigned });

  it("reads user consent as off when only owned-resource entries are assigned, and the workflow follows", async () => {
    const { result } = await run({
      ...consent([OWNED_TEAMS, OWNED_CHAT]),
      workflow: { isEnabled: false },
    });
    expect(status(result, "user-consent")).toBe("Meets");
    expect(value(result, "user-consent")).toMatch(/^Off\. Users cannot consent to apps\. /);
    expect(value(result, "user-consent")).toContain(
      "Group owners can consent for their groups and teams: ",
    );
    expect(value(result, "user-consent")).toContain(OWNED_TEAMS);
    expect(status(result, "admin-consent-workflow")).toBe("Review");
    expect(cell(result, "admin-consent-workflow", "expected")).toBe("On");
  });

  it("reads user consent as limited with the low-impact policy plus an owned-resource entry", async () => {
    const { result } = await run({
      ...consent([LOW, OWNED_TEAMS]),
      workflow: { isEnabled: false },
    });
    expect(status(result, "user-consent")).toBe("Meets");
    expect(value(result, "user-consent")).toContain("On: Low impact from verified publishers");
    expect(value(result, "user-consent")).toContain("Group owners can consent");
    expect(status(result, "admin-consent-workflow")).toBe("Review");
  });

  it("still needs review when a ForSelf entry beyond the low-impact one is present", async () => {
    const { result } = await run({
      ...consent([LEGACY, OWNED_TEAMS]),
      workflow: { isEnabled: false },
    });
    expect(status(result, "user-consent")).toBe("Review");
    expect(status(result, "admin-consent-workflow")).toBe("Meets");
  });

  it("is Not read when the list holds anything but text", async () => {
    for (const assigned of [[null], [LOW, 3], [{ id: LOW }], [[LOW]]]) {
      const { result } = await run(consent(assigned));
      expect(status(result, "user-consent"), JSON.stringify(assigned)).toBe("Not read");
    }
  });
});

describe("a missing field is Not read, never Meets", () => {
  it("shows every row as Not read for empty answers, without calling the run partial", async () => {
    const { result } = await run({
      defaults: {},
      methods: {},
      authorization: {},
      workflow: {},
      grants: [],
    });
    expect(result.rows).toHaveLength(22);
    for (const item of result.rows) {
      expect(item.cells.status, item.key).toBe("Not read");
      expect(item.cells.value, item.key).toBe("Microsoft did not return this setting");
    }
    expect(result.summary.map((tile) => tile.value)).toEqual([0, 0, 22]);
    expect(result.complete).toBe(true);
    expect(result.notes.join("\n")).toContain("22 settings were missing");
  });

  it("is Not read for a field of the wrong type", async () => {
    const body = withAuthorization(
      healthy(),
      { allowEmailVerifiedUsersToJoinOrganization: "false", allowInvitesFrom: 3 },
      { allowedToCreateApps: "no", permissionGrantPoliciesAssigned: "none" },
    );
    const { result } = await run(body);
    for (const key of ["email-verified-join", "guest-invites", "create-apps", "user-consent"])
      expect(status(result, key), key).toBe("Not read");
  });
});

describe("a path that fails", () => {
  const cases: [string, string, string[]][] = [
    [DEFAULTS, "security defaults policy", [SECURITY_DEFAULTS_KEY]],
    [METHODS, "authentication methods policy", ["method-sms", "registration-campaign"]],
    [AUTHORIZATION, "authorization policy", ["guest-invites", "user-consent", "msol-powershell"]],
    [WORKFLOW, "admin consent request policy", ["admin-consent-workflow"]],
  ];

  for (const [path, name, keys] of cases) {
    it(`leaves the ${name} rows Not read, says why, and makes the run partial`, async () => {
      const { result } = await run(healthy(), { failing: { [path]: 403 } });
      for (const key of keys) {
        expect(status(result, key), key).toBe("Not read");
        expect(value(result, key), key).toBe("Microsoft refused (HTTP 403)");
      }
      expect(result.complete).toBe(false);
      const note = result.notes.find((item) => item.includes(name));
      expect(note).toContain("could not be read: Microsoft refused (HTTP 403)");
      expect(note).toContain("show Not read");
      // Everything from the other policies is still read.
      const others = result.rows.filter((item) => item.cells.status === "Meets");
      expect(others.length).toBeGreaterThan(0);
    });
  }

  it("uses plain wording for a missing policy, a server error and throttling", async () => {
    const missing = await run(healthy(), { failing: { [WORKFLOW]: 404 } });
    expect(value(missing.result, "admin-consent-workflow")).toBe(
      "Microsoft has no such policy (HTTP 404)",
    );
    const broken = await run(healthy(), { failing: { [WORKFLOW]: 500 } });
    expect(value(broken.result, "admin-consent-workflow")).toBe("Microsoft answered HTTP 500");
    const limited = await run(healthy(), { failing: { [WORKFLOW]: 429 } });
    expect(value(limited.result, "admin-consent-workflow")).toBe(
      "Microsoft is limiting requests, so run the audit again in a few minutes",
    );
  });

  it("does not turn a refused session into a partial result", async () => {
    await expect(run(healthy(), { failing: { [AUTHORIZATION]: 401 } })).rejects.toBeInstanceOf(
      GraphRequestError,
    );
  });

  it("survives every path failing, with all rows Not read", async () => {
    const failing = Object.fromEntries(PATHS.map((path) => [path, 403]));
    const { result } = await run(healthy(), { failing });
    expect(result.rows).toHaveLength(22);
    for (const item of result.rows) expect(item.cells.status, item.key).toBe("Not read");
    expect(result.complete).toBe(false);
  });
});

describe("the authorization policy answered as a list", () => {
  it("reads the first entry when Microsoft wraps it in a collection", async () => {
    const network = fake(tenantSettingsAudit, (url) => {
      if (url.pathname === AUTHORIZATION)
        return Response.json({
          value: [withAuthorization(healthy(), { allowInvitesFrom: "none" }).authorization],
        });
      return routeOf(healthy())(url);
    });
    const result = await tenantSettingsAudit.run(network.context);
    expect(value(result, "guest-invites")).toBe("No one");
  });
});

describe("ordering and wording", () => {
  it("lists Review first, then Not read, then Meets, each in area order", async () => {
    const body = withAuthorization(
      withMethod(healthy(), "Sms", "enabled"),
      { allowInvitesFrom: "everyone" },
      { allowedToCreateApps: true },
    );
    const { result } = await run({ ...body, workflow: { isEnabled: false }, defaults: {} });
    const states = result.rows.map((item) => item.cells.status);
    const sorted = [...states].sort(
      (a, b) =>
        ["Review", "Not read", "Meets"].indexOf(String(a)) -
        ["Review", "Not read", "Meets"].indexOf(String(b)),
    );
    expect(states).toEqual(sorted);
    expect(result.rows[0]?.cells.status).toBe("Review");
  });

  it("keeps the audit's own text free of em dashes, middle dots and exclamation marks", async () => {
    const { result } = await run(
      off(withAuthorization(healthy(), { allowInvitesFrom: "everyone" })),
      { run: { rows: [caRow()] } },
    );
    const text = JSON.stringify([result.rows, result.notes, auditDefinition(tenantSettingsAudit)]);
    expect(text).not.toMatch(/[—–·!]/);
    for (const note of result.notes) expect(note.length).toBeLessThanOrEqual(500);
  });

  it("ends with what it cannot see", async () => {
    const { result } = await run(healthy());
    expect(result.notes.at(-1)).toContain("per-user authentication method settings");
    expect(result.notes.at(-1)).toContain("self-service password reset configuration");
    expect(result.notes.at(-1)).toContain("Conditional Access coverage");
  });
});
