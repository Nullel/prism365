import { auditRowSchema, queryAuditRows } from "@prism365/contracts";
import { GraphHttpClient, GraphRequestError } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  type Audit,
  type AuditContext,
  activityOf,
  auditDefinition,
  auditReadScope,
  builtInAudits,
  type DirectoryUser,
  domainCounts,
  guestAccountsAudit,
  guestDomain,
  inactiveAccountsAudit,
  licenseCapacityAudit,
  licenseReclaimAudit,
  mfaRegistrationAudit,
  principalDomain,
  scopeNote,
  strongestMethod,
  tagOf,
} from "../src/index.ts";

// Synthetic tenant data only: invented names, example.invalid domains, made-up SKU IDs.
const now = new Date("2026-10-22T12:00:00Z");
const daysAgo = (days: number) => new Date(now.getTime() - days * 86_400_000).toISOString();
const E3 = "11111111-0000-4000-8000-0000000000e3";
const BASIC = "11111111-0000-4000-8000-0000000000b1";

const users = [
  {
    id: "u-active",
    displayName: "Avery Active",
    userPrincipalName: "avery@example.invalid",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(900),
    assignedLicenses: [{ skuId: E3 }],
    licenseAssignmentStates: [{ skuId: E3, assignedByGroup: null }],
    signInActivity: {
      lastSuccessfulSignInDateTime: daysAgo(2),
      lastSignInDateTime: daysAgo(2),
      lastNonInteractiveSignInDateTime: daysAgo(1),
    },
  },
  {
    id: "u-stale",
    displayName: "Sam Stale",
    userPrincipalName: "sam@example.invalid",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(800),
    assignedLicenses: [{ skuId: E3 }, { skuId: BASIC }],
    licenseAssignmentStates: [
      { skuId: E3, assignedByGroup: "group-1" },
      { skuId: BASIC, assignedByGroup: null },
    ],
    signInActivity: {
      lastSuccessfulSignInDateTime: daysAgo(130),
      lastSignInDateTime: daysAgo(130),
      lastNonInteractiveSignInDateTime: null,
    },
  },
  {
    id: "u-disabled",
    displayName: "Dana Departed",
    userPrincipalName: "dana@example.invalid",
    userType: "Member",
    accountEnabled: false,
    createdDateTime: daysAgo(700),
    assignedLicenses: [{ skuId: BASIC }],
    licenseAssignmentStates: [{ skuId: BASIC, assignedByGroup: null }],
    signInActivity: {
      lastSuccessfulSignInDateTime: daysAgo(10),
      lastSignInDateTime: daysAgo(10),
      lastNonInteractiveSignInDateTime: null,
    },
  },
  {
    id: "u-never",
    displayName: "Nico New",
    userPrincipalName: "nico@example.invalid",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(45),
    assignedLicenses: [],
    signInActivity: null,
  },
  {
    id: "u-failing",
    displayName: "Fran Failing",
    userPrincipalName: "fran@example.invalid",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(600),
    assignedLicenses: [],
    signInActivity: {
      lastSuccessfulSignInDateTime: null,
      lastSignInDateTime: daysAgo(3),
      lastNonInteractiveSignInDateTime: null,
    },
  },
  {
    id: "g-pending",
    displayName: "Pat Partner",
    userPrincipalName: "pat_partner.example.invalid#EXT#@tenant.example.invalid",
    mail: null,
    userType: "Guest",
    accountEnabled: true,
    createdDateTime: daysAgo(40),
    externalUserState: "PendingAcceptance",
    assignedLicenses: [],
    signInActivity: null,
  },
  {
    id: "g-quiet",
    displayName: '=HYPERLINK("http://example.invalid")',
    userPrincipalName: "quinn_vendor.example.invalid#EXT#@tenant.example.invalid",
    mail: "quinn@vendor.example.invalid",
    userType: "Guest",
    accountEnabled: true,
    createdDateTime: daysAgo(400),
    externalUserState: "Accepted",
    assignedLicenses: [],
    signInActivity: {
      lastSuccessfulSignInDateTime: daysAgo(200),
      lastSignInDateTime: daysAgo(200),
      lastNonInteractiveSignInDateTime: null,
    },
  },
];

const skus = [
  {
    skuId: E3,
    skuPartNumber: "SPE_E3",
    consumedUnits: 12,
    prepaidUnits: { enabled: 10, suspended: 0, warning: 0, lockedOut: 0 },
    capabilityStatus: "Enabled",
    appliesTo: "User",
  },
  {
    skuId: BASIC,
    skuPartNumber: "O365_BUSINESS_ESSENTIALS",
    consumedUnits: 3,
    prepaidUnits: { enabled: 8, suspended: 0, warning: 0, lockedOut: 0 },
    capabilityStatus: "Enabled",
    appliesTo: "User",
  },
  {
    skuId: "11111111-0000-4000-8000-00000000f1ee",
    skuPartNumber: "FLOW_FREE",
    consumedUnits: 4,
    prepaidUnits: { enabled: 10_000, suspended: 0, warning: 0, lockedOut: 0 },
    capabilityStatus: "Enabled",
    appliesTo: "User",
  },
  {
    skuId: "11111111-0000-4000-8000-00000000dead",
    skuPartNumber: "VISIOCLIENT",
    consumedUnits: 1,
    prepaidUnits: { enabled: 0, suspended: 0, warning: 2, lockedOut: 0 },
    capabilityStatus: "Warning",
    appliesTo: "User",
  },
];

const registrations = [
  {
    id: "u-active",
    userPrincipalName: "avery@example.invalid",
    userDisplayName: "Avery Active",
    isAdmin: true,
    isMfaRegistered: true,
    isMfaCapable: true,
    isPasswordlessCapable: false,
    methodsRegistered: ["mobilePhone"],
    userPreferredMethodForSecondaryAuthentication: "sms",
    userType: "member",
  },
  {
    id: "u-stale",
    userPrincipalName: "sam@example.invalid",
    userDisplayName: "Sam Stale",
    isAdmin: true,
    isMfaRegistered: false,
    methodsRegistered: [],
    userType: "member",
  },
  {
    id: "u-never",
    userPrincipalName: "nico@example.invalid",
    userDisplayName: "Nico New",
    isAdmin: false,
    isMfaRegistered: true,
    isPasswordlessCapable: true,
    methodsRegistered: ["microsoftAuthenticatorPush", "fido2SecurityKey"],
    isSystemPreferredAuthenticationMethodEnabled: true,
    systemPreferredAuthenticationMethods: ["Fido2"],
    userType: "member",
  },
  // Unreadable: no id.
  { userPrincipalName: "broken@example.invalid" },
];

// One policy and one risky user, so the generic definition checks see a row from each audit.
const policies = [
  {
    id: "p-block-legacy",
    displayName: "Block legacy authentication",
    state: "enabled",
    conditions: {
      users: { includeUsers: ["All"], excludeUsers: [] },
      applications: { includeApplications: ["All"] },
      clientAppTypes: ["exchangeActiveSync", "other"],
    },
    grantControls: { operator: "OR", builtInControls: ["block"] },
  },
];
const riskyUsers = [
  {
    id: "u-stale",
    isDeleted: false,
    isProcessing: false,
    riskLastUpdatedDateTime: daysAgo(3),
    riskLevel: "high",
    riskState: "atRisk",
    riskDetail: "none",
    userDisplayName: "Sam Stale",
    userPrincipalName: "sam@example.invalid",
  },
];

type Route = (url: URL) => Response | undefined;

function context(
  audit: Audit,
  route: Route,
  granted = ["User.Read.All", "AuditLog.Read.All", "LicenseAssignment.Read.All"],
  scope: ReadonlySet<string> | null = null,
): AuditContext & { requests: URL[] } {
  const requests: URL[] = [];
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    maxResponseBytes: 8_388_608,
    fetchImplementation: async (input) => {
      const url = new URL(String(input));
      requests.push(url);
      return route(url) ?? new Response("{}", { status: 404 });
    },
  });
  return {
    graph: scopeGraph(client, [...audit.packs, ...audit.optionalPacks]),
    now,
    granted: new Set(granted),
    signal: new AbortController().signal,
    scope,
    requests,
  };
}

// Returns only the fields a request selected, as Graph does.
function selected(url: URL, records: Record<string, unknown>[]) {
  const fields = url.searchParams.get("$select")?.split(",");
  if (!fields) return records;
  return records.map((record) =>
    Object.fromEntries(Object.entries(record).filter(([key]) => fields.includes(key))),
  );
}

const tenant: Route = (url) => {
  if (url.pathname === "/v1.0/users") return Response.json({ value: selected(url, users) });
  if (url.pathname === "/v1.0/subscribedSkus") return Response.json({ value: skus });
  if (url.pathname === "/v1.0/reports/authenticationMethods/userRegistrationDetails")
    return Response.json({ value: registrations });
  if (url.pathname === "/v1.0/identity/conditionalAccess/policies")
    return Response.json({ value: policies });
  if (url.pathname === "/v1.0/identityProtection/riskyUsers")
    return Response.json({ value: riskyUsers });
  return undefined;
};

const row = (rows: { key: string }[], key: string) => {
  const found = rows.find((item) => item.key === key);
  if (!found) throw new Error(`No row ${key}`);
  return found as (typeof rows)[number] & {
    tags: string[];
    cells: Record<string, unknown>;
  };
};

describe("audit definitions", () => {
  it("publish valid tables whose views, sort and rows agree", async () => {
    for (const audit of builtInAudits) {
      const definition = auditDefinition(audit);
      const keys = new Set(definition.columns.map((column) => column.key));
      expect(keys.has(definition.defaultSort.column), definition.id).toBe(true);
      expect(definition.views.map((view) => view.id)).toContain(definition.defaultView);
      expect(definition.permissions.length, definition.id).toBeGreaterThan(0);
      const result = await audit.run(context(audit, tenant));
      for (const item of result.rows) {
        auditRowSchema.parse(item);
        for (const key of Object.keys(item.cells))
          expect(keys.has(key), `${definition.id}.${key}`).toBe(true);
      }
      for (const tile of result.summary)
        if (tile.view) expect(definition.views.map((view) => view.id)).toContain(tile.view);
    }
  });

  it("derive permissions from read packs", () => {
    expect(auditDefinition(inactiveAccountsAudit)).toMatchObject({
      permissions: ["User.Read.All", "AuditLog.Read.All"],
      optionalPermissions: ["LicenseAssignment.Read.All"],
    });
    expect(auditDefinition(mfaRegistrationAudit).permissions).toEqual(["AuditLog.Read.All"]);
  });
});

describe("how inactivity is counted", () => {
  const user = (activity: DirectoryUser["signInActivity"], created = daysAgo(1_000)) =>
    ({ createdDateTime: created, signInActivity: activity }) as DirectoryUser;

  it("counts from the last successful sign-in", () => {
    expect(
      activityOf(
        user({
          lastSuccessfulSignInDateTime: daysAgo(40),
          lastSignInDateTime: daysAgo(1),
          lastNonInteractiveSignInDateTime: null,
        }),
        now,
      ),
    ).toMatchObject({ basis: "success", days: 40, noSuccess: false });
  });

  it("falls back to the last attempt, flagging attempts with no success since December 2023", () => {
    const recent = activityOf(
      user({
        lastSuccessfulSignInDateTime: null,
        lastSignInDateTime: daysAgo(5),
        lastNonInteractiveSignInDateTime: daysAgo(9),
      }),
      now,
    );
    expect(recent).toMatchObject({ basis: "attempt", days: 5, noSuccess: true });
    const old = activityOf(
      user({
        lastSuccessfulSignInDateTime: null,
        lastSignInDateTime: "2022-05-01T00:00:00Z",
        lastNonInteractiveSignInDateTime: null,
      }),
      now,
    );
    expect(old).toMatchObject({ basis: "attempt", noSuccess: false });
  });

  it("counts a never-used account from creation, but not from before Microsoft's records", () => {
    expect(activityOf(user(null, daysAgo(45)), now)).toMatchObject({ basis: "none", days: 45 });
    const ancient = activityOf(user(null, "2015-01-01T00:00:00Z"), now);
    expect(new Date(ancient.from ?? 0).toISOString()).toBe("2020-04-01T00:00:00.000Z");
  });

  it("finds a guest's home domain from its email or its external principal name", () => {
    expect(guestDomain(users[5] as unknown as DirectoryUser)).toBe("partner.example.invalid");
    expect(guestDomain(users[6] as unknown as DirectoryUser)).toBe("vendor.example.invalid");
  });

  it("ranks sign-in methods", () => {
    expect(strongestMethod(["mobilePhone"], false)).toBe("Phone or text");
    expect(strongestMethod(["mobilePhone", "softwareOneTimePasscode"], false)).toBe("App or token");
    expect(strongestMethod(["email"], true)).toBe("Passwordless");
    expect(strongestMethod(["email", "securityQuestion"], false)).toBe("None");
  });
});

describe("row tags", () => {
  it("turns provider values into short lowercase slugs", () => {
    expect(tagOf("Mobile  Phone (SMS)!")).toEqual(["mobile-phone-sms"]);
    expect(tagOf("--Member--")).toEqual(["member"]);
    expect(tagOf(`${"a".repeat(39)} b`)).toEqual(["a".repeat(39)]);
    expect(tagOf(null)).toEqual([]);
    expect(tagOf("!!! ---")).toEqual([]);
  });

  it("handles long runs of separators in one pass", () => {
    const value = `${"-".repeat(100_000)}x${" -".repeat(100_000)}`;
    const started = performance.now();
    expect(tagOf(value)).toEqual(["x"]);
    expect(performance.now() - started).toBeLessThan(500);
  });
});

describe("inactive accounts", () => {
  it("buckets every account by days and names licenses", async () => {
    const ctx = context(inactiveAccountsAudit, tenant);
    const result = await inactiveAccountsAudit.run(ctx);
    expect(result.rows).toHaveLength(users.length);
    expect(row(result.rows, "u-stale")).toMatchObject({
      tags: expect.arrayContaining(["d7", "d30", "d60", "d90", "d120", "enabled", "licensed"]),
      cells: {
        days: 130,
        basis: "Successful sign-in",
        licenses: ["Microsoft 365 Business Basic", "Microsoft 365 E3"],
      },
    });
    expect(row(result.rows, "u-stale").tags).not.toContain("d180");
    expect(row(result.rows, "u-never")).toMatchObject({
      tags: expect.arrayContaining(["never", "d30"]),
      cells: { days: 45, basis: "Account created" },
    });
    expect(row(result.rows, "u-failing").tags).toContain("no-success");
    const definition = auditDefinition(inactiveAccountsAudit);
    const ninety = queryAuditRows(definition, result.rows, { view: "d90" });
    expect(ninety.map((item) => item.key)).toEqual(["g-quiet", "u-stale"]);
    expect(result.summary.find((tile) => tile.view === "d30")?.value).toBe(4);
    // One page of up to 500 users with sign-in dates, as Graph requires.
    const read = ctx.requests.find((url) => url.pathname === "/v1.0/users");
    expect(read?.searchParams.get("$top")).toBe("500");
    expect(read?.searchParams.get("$select")).toContain("signInActivity");
    expect(result.complete).toBe(true);
  });

  it("fails with Microsoft's refusal when sign-in dates are not allowed", async () => {
    const refused: Route = (url) =>
      url.searchParams.get("$select")?.includes("signInActivity")
        ? new Response("{}", { status: 403 })
        : tenant(url);
    await expect(
      inactiveAccountsAudit.run(context(inactiveAccountsAudit, refused)),
    ).rejects.toThrow(GraphRequestError);
  });

  it("leaves out an unreadable account and says the result is partial", async () => {
    const broken: Route = (url) =>
      url.pathname === "/v1.0/users"
        ? Response.json({ value: [...users, { displayName: "no id" }] })
        : tenant(url);
    const result = await inactiveAccountsAudit.run(context(inactiveAccountsAudit, broken));
    expect(result.rows).toHaveLength(users.length);
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("1 account could not be read");
  });

  it("shows product IDs when license names cannot be read", async () => {
    const result = await inactiveAccountsAudit.run(
      context(inactiveAccountsAudit, tenant, ["User.Read.All", "AuditLog.Read.All"]),
    );
    expect(row(result.rows, "u-disabled").cells.licenses).toEqual([BASIC]);
    expect(result.notes.join(" ")).toContain("product IDs");
  });
});

describe("dates Microsoft could not have recorded", () => {
  const route =
    (extra: Record<string, unknown>[]): Route =>
    (url) =>
      url.pathname === "/v1.0/users"
        ? Response.json({ value: selected(url, [...users, ...extra]) })
        : tenant(url);
  const active = {
    displayName: "Active Recently",
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(300),
    assignedLicenses: [{ skuId: E3 }],
  };

  it("never counts an account with an unreadable or impossible date as inactive", async () => {
    const result = await inactiveAccountsAudit.run(
      context(
        inactiveAccountsAudit,
        route([
          {
            ...active,
            id: "u-sentinel",
            userPrincipalName: "sentinel@example.invalid",
            signInActivity: {
              lastSuccessfulSignInDateTime: "0001-01-01T00:00:00Z",
              lastSignInDateTime: daysAgo(1),
            },
          },
          {
            ...active,
            id: "u-garbled",
            userPrincipalName: "garbled@example.invalid",
            signInActivity: { lastSuccessfulSignInDateTime: 12345, lastSignInDateTime: daysAgo(1) },
          },
          {
            ...active,
            id: "u-future",
            userPrincipalName: "future@example.invalid",
            signInActivity: { lastSignInDateTime: "2099-01-01T00:00:00Z" },
          },
          {
            ...active,
            id: "u-object",
            userPrincipalName: "object@example.invalid",
            signInActivity: "not an object",
          },
        ]),
      ),
    );
    for (const key of ["u-sentinel", "u-garbled", "u-future", "u-object"]) {
      const item = row(result.rows, key);
      expect(item.tags, key).toEqual(expect.arrayContaining(["unreadable"]));
      expect(item.tags, key).not.toContain("never");
      expect(
        item.tags.some((tag) => /^d\d+$/.test(tag)),
        key,
      ).toBe(false);
      expect(item.cells, key).toMatchObject({ days: null, basis: "Unreadable dates" });
    }
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain(
      "4 accounts have sign-in dates Prism365 could not read",
    );
    // They are not reclaim candidates either.
    const reclaim = await licenseReclaimAudit.run(
      context(
        licenseReclaimAudit,
        route([
          {
            ...active,
            id: "u-sentinel",
            userPrincipalName: "sentinel@example.invalid",
            signInActivity: { lastSuccessfulSignInDateTime: "0001-01-01T00:00:00Z" },
          },
        ]),
      ),
    );
    expect(reclaim.rows.map((item) => item.key)).not.toContain("u-sentinel");
    expect(reclaim.complete).toBe(false);
  });

  it("removes characters storage refuses from provider text", async () => {
    const result = await inactiveAccountsAudit.run(
      context(
        inactiveAccountsAudit,
        route([
          {
            ...active,
            id: "u-nul",
            displayName: "Nul\u0000Name \ud800",
            userPrincipalName: "nul@example.invalid",
            signInActivity: { lastSuccessfulSignInDateTime: daysAgo(3) },
          },
        ]),
      ),
    );
    expect(row(result.rows, "u-nul").cells.person).toEqual({
      name: "NulName \uFFFD",
      principal: "nul@example.invalid",
    });
  });
});

describe("licenses to reclaim", () => {
  it("lists licensed accounts that are disabled or 30+ days inactive, with group assignment", async () => {
    const result = await licenseReclaimAudit.run(context(licenseReclaimAudit, tenant));
    expect(result.rows.map((item) => item.key).sort()).toEqual(["u-disabled", "u-stale"]);
    expect(row(result.rows, "u-stale")).toMatchObject({
      tags: expect.arrayContaining(["group", "d90"]),
      cells: { assignment: "Directly and by group", licenseCount: 2 },
    });
    expect(row(result.rows, "u-disabled").cells).toMatchObject({
      status: "Disabled",
      assignment: "Directly",
    });
    expect(result.summary.find((tile) => tile.label === "Licenses held")?.value).toBe(3);
    expect(result.notes[0]).toContain("By product:");
  });

  it("lists only disabled accounts when Microsoft refuses sign-in dates", async () => {
    const refused: Route = (url) =>
      url.searchParams.get("$select")?.includes("signInActivity")
        ? new Response("{}", { status: 403 })
        : tenant(url);
    const result = await licenseReclaimAudit.run(context(licenseReclaimAudit, refused));
    expect(result.rows.map((item) => item.key)).toEqual(["u-disabled"]);
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("Microsoft refused last sign-in dates");
  });

  it("drops group assignment when the tenant rejects that field", async () => {
    const noStates: Route = (url) =>
      url.searchParams.get("$select")?.includes("licenseAssignmentStates")
        ? new Response("{}", { status: 400 })
        : tenant(url);
    const result = await licenseReclaimAudit.run(context(licenseReclaimAudit, noStates));
    expect(row(result.rows, "u-stale").cells.assignment).toBe("Unknown");
  });
});

describe("guest accounts", () => {
  it("finds invitations not accepted and guests gone quiet", async () => {
    const result = await guestAccountsAudit.run(context(guestAccountsAudit, tenant));
    expect(result.rows.map((item) => item.key).sort()).toEqual(["g-pending", "g-quiet"]);
    expect(row(result.rows, "g-pending")).toMatchObject({
      tags: expect.arrayContaining(["pending", "pending-30", "never"]),
      cells: { invitation: "Not accepted", domain: "partner.example.invalid" },
    });
    expect(row(result.rows, "g-quiet").tags).toEqual(expect.arrayContaining(["d90", "d180"]));
  });
});

describe("MFA registration", () => {
  it("flags admins without MFA and phone-only methods", async () => {
    const result = await mfaRegistrationAudit.run(context(mfaRegistrationAudit, tenant));
    expect(result.rows).toHaveLength(3);
    expect(row(result.rows, "u-stale").tags).toEqual(expect.arrayContaining(["admin", "no-mfa"]));
    expect(row(result.rows, "u-active")).toMatchObject({
      tags: expect.arrayContaining(["admin", "phone-only"]),
      cells: { strength: "Phone or text", preferred: "Text message" },
    });
    expect(row(result.rows, "u-never")).toMatchObject({
      tags: expect.arrayContaining(["passwordless"]),
      cells: {
        methods: ["Authenticator app", "Security key (FIDO2)"],
        preferred: "Security key (FIDO2) (chosen by Microsoft)",
      },
    });
    expect(result.summary.find((tile) => tile.view === "admins-no-mfa")).toMatchObject({
      value: 1,
      tone: "critical",
    });
    expect(result.complete).toBe(false);
  });
});

describe("license capacity", () => {
  it("flags products with no seats left, over-assigned or expired, and ignores free seats", async () => {
    const result = await licenseCapacityAudit.run(context(licenseCapacityAudit, tenant));
    expect(row(result.rows, E3)).toMatchObject({
      tags: expect.arrayContaining(["attention", "full"]),
      cells: { product: "Microsoft 365 E3", available: -2 },
    });
    expect(row(result.rows, BASIC).tags).toEqual(["spare", "user"]);
    expect(row(result.rows, "11111111-0000-4000-8000-00000000f1ee").tags).not.toContain("spare");
    expect(row(result.rows, "11111111-0000-4000-8000-00000000dead")).toMatchObject({
      tags: expect.arrayContaining(["attention", "expiring"]),
      cells: { status: "Expired, in grace period", grace: 2, available: 1 },
    });
    expect(result.summary.find((tile) => tile.view === "spare")?.value).toBe(5);
  });
});

// Account audits limited to the domains an administrator chose. Two organisations' domains, a
// mixed-case principal name, an account with no principal name, and a guest.
describe("audit scope", () => {
  const scopeUsers = [
    {
      id: "s-a1",
      displayName: "Ada Alpha",
      userPrincipalName: "ada@a.example",
      userType: "Member",
      accountEnabled: true,
      createdDateTime: daysAgo(900),
      assignedLicenses: [{ skuId: E3 }],
      licenseAssignmentStates: [{ skuId: E3, assignedByGroup: null }],
      signInActivity: {
        lastSuccessfulSignInDateTime: daysAgo(200),
        lastSignInDateTime: daysAgo(200),
        lastNonInteractiveSignInDateTime: null,
      },
    },
    {
      id: "s-a2",
      displayName: "Abe Alpha",
      userPrincipalName: "abe@A.Example",
      userType: "Member",
      accountEnabled: true,
      createdDateTime: daysAgo(900),
      assignedLicenses: [{ skuId: E3 }],
      licenseAssignmentStates: [{ skuId: E3, assignedByGroup: null }],
      signInActivity: {
        lastSuccessfulSignInDateTime: daysAgo(3),
        lastSignInDateTime: daysAgo(3),
        lastNonInteractiveSignInDateTime: null,
      },
    },
    {
      id: "s-b1",
      displayName: "Bo Beta",
      userPrincipalName: "bo@b.example",
      userType: "Member",
      accountEnabled: false,
      createdDateTime: daysAgo(900),
      assignedLicenses: [{ skuId: BASIC }],
      licenseAssignmentStates: [{ skuId: BASIC, assignedByGroup: null }],
      signInActivity: {
        lastSuccessfulSignInDateTime: daysAgo(10),
        lastSignInDateTime: daysAgo(10),
        lastNonInteractiveSignInDateTime: null,
      },
    },
    {
      id: "s-b2",
      displayName: "Bea Beta",
      userPrincipalName: "bea@b.example",
      userType: "Member",
      accountEnabled: true,
      createdDateTime: daysAgo(900),
      assignedLicenses: [{ skuId: E3 }],
      licenseAssignmentStates: [{ skuId: E3, assignedByGroup: null }],
      signInActivity: {
        lastSuccessfulSignInDateTime: daysAgo(150),
        lastSignInDateTime: daysAgo(150),
        lastNonInteractiveSignInDateTime: null,
      },
    },
    {
      id: "s-c1",
      displayName: "Cy Gamma",
      userPrincipalName: "cy@c.example",
      userType: "Member",
      accountEnabled: true,
      createdDateTime: daysAgo(100),
      assignedLicenses: [],
      signInActivity: null,
    },
    {
      id: "s-none",
      displayName: "Nan Nameless",
      userPrincipalName: null,
      userType: "Member",
      accountEnabled: true,
      createdDateTime: daysAgo(900),
      assignedLicenses: [{ skuId: E3 }],
      licenseAssignmentStates: [{ skuId: E3, assignedByGroup: null }],
      signInActivity: {
        lastSuccessfulSignInDateTime: daysAgo(250),
        lastSignInDateTime: daysAgo(250),
        lastNonInteractiveSignInDateTime: null,
      },
    },
    {
      id: "s-guest",
      displayName: "Gus Guest",
      userPrincipalName: "gus_vendor.example#EXT#@tenant.example",
      mail: "gus@vendor.example",
      userType: "Guest",
      accountEnabled: true,
      createdDateTime: daysAgo(400),
      externalUserState: "PendingAcceptance",
      assignedLicenses: [],
      signInActivity: null,
    },
  ];
  const scopeRegistrations = [
    {
      id: "s-a1",
      userPrincipalName: "ada@a.example",
      userDisplayName: "Ada Alpha",
      isAdmin: true,
      isMfaRegistered: false,
      methodsRegistered: [],
      userType: "member",
    },
    {
      id: "s-a2",
      userPrincipalName: "abe@A.Example",
      userDisplayName: "Abe Alpha",
      isAdmin: false,
      isMfaRegistered: true,
      methodsRegistered: ["mobilePhone"],
      userType: "member",
    },
    {
      id: "s-b2",
      userPrincipalName: "bea@b.example",
      userDisplayName: "Bea Beta",
      isAdmin: true,
      isMfaRegistered: false,
      methodsRegistered: [],
      userType: "member",
    },
    {
      id: "s-c1",
      userPrincipalName: "cy@c.example",
      userDisplayName: "Cy Gamma",
      isAdmin: false,
      isMfaRegistered: false,
      methodsRegistered: [],
      userType: "member",
    },
    {
      id: "s-none",
      userPrincipalName: null,
      userDisplayName: "Nan Nameless",
      isAdmin: false,
      isMfaRegistered: false,
      methodsRegistered: [],
      userType: "member",
    },
  ];
  const scoped: Route = (url) => {
    if (url.pathname === "/v1.0/users") return Response.json({ value: selected(url, scopeUsers) });
    if (url.pathname === "/v1.0/reports/authenticationMethods/userRegistrationDetails")
      return Response.json({ value: scopeRegistrations });
    return tenant(url);
  };
  const only = (...domains: string[]) => new Set(domains);
  const run = (audit: Audit, scope: ReadonlySet<string> | null) =>
    audit.run(context(audit, scoped, undefined, scope));
  const keys = (result: { rows: { key: string }[] }) => result.rows.map((item) => item.key).sort();
  const tile = (result: { summary: { label: string; value: number }[] }, label: string) =>
    result.summary.find((item) => item.label === label)?.value;

  it("publishes which audits use the scope, and a Domain filter on each of them", () => {
    for (const audit of [inactiveAccountsAudit, licenseReclaimAudit, mfaRegistrationAudit]) {
      const definition = auditDefinition(audit);
      expect(definition.scoped, definition.id).toBe(true);
      expect(definition.columns.find((column) => column.key === "domain")).toMatchObject({
        label: "Domain",
        kind: "enum",
        facet: true,
      });
    }
    for (const audit of [guestAccountsAudit, licenseCapacityAudit])
      expect(auditDefinition(audit).scoped, audit.definition.id).toBe(false);
  });

  it("reads the domain after the last @, in lower case", () => {
    expect(principalDomain("Ada@A.Example")).toBe("a.example");
    expect(principalDomain("odd@name@c.example")).toBe("c.example");
    expect(principalDomain("gus_vendor.example#EXT#@tenant.example")).toBe("tenant.example");
    expect(principalDomain("no-at-sign")).toBeNull();
    expect(principalDomain("trailing@")).toBeNull();
    expect(principalDomain("")).toBeNull();
    expect(principalDomain(null)).toBeNull();
  });

  it("limits the inactive accounts audit before it counts anything", async () => {
    const everyone = await run(inactiveAccountsAudit, null);
    expect(keys(everyone)).toEqual(
      ["s-a1", "s-a2", "s-b1", "s-b2", "s-c1", "s-guest", "s-none"].sort(),
    );
    const result = await run(inactiveAccountsAudit, only("a.example"));
    expect(keys(result)).toEqual(["s-a1", "s-a2"]);
    expect(tile(result, "Accounts")).toBe(2);
    // Ada's 200 days is the only one in scope: the unscoped run counts five such accounts.
    expect(tile(everyone, "Inactive 30+ days")).toBe(5);
    expect(tile(result, "Inactive 30+ days")).toBe(1);
    expect(tile(result, "Never signed in")).toBe(0);
    expect(result.notes[0]).toBe("Limited to accounts in: a.example.");
    expect(row(result.rows, "s-a2").cells.domain).toBe("a.example");
    expect(result.complete).toBe(true);
  });

  it("names every domain in the note, in order", async () => {
    const result = await run(inactiveAccountsAudit, only("b.example", "a.example"));
    expect(keys(result)).toEqual(["s-a1", "s-a2", "s-b1", "s-b2"]);
    expect(result.notes[0]).toBe("Limited to accounts in: a.example, b.example.");
  });

  it("keeps an account with no principal name only when every domain is in scope", async () => {
    expect(keys(await run(inactiveAccountsAudit, only("c.example")))).toEqual(["s-c1"]);
    for (const everything of [null, new Set<string>()]) {
      const result = await run(inactiveAccountsAudit, everything);
      expect(keys(result)).toContain("s-none");
      expect(row(result.rows, "s-none").cells.domain).toBeNull();
      expect(result.notes.join(" ")).not.toContain("Limited to");
    }
  });

  it("gives guests their tenant's domain in the Domain column", async () => {
    const result = await run(inactiveAccountsAudit, null);
    expect(row(result.rows, "s-guest").cells.domain).toBe("tenant.example");
    expect(keys(await run(inactiveAccountsAudit, only("tenant.example")))).toEqual(["s-guest"]);
  });

  it("limits licenses to reclaim, its tiles and its product note", async () => {
    const everyone = await run(licenseReclaimAudit, null);
    expect(keys(everyone)).toEqual(["s-a1", "s-b1", "s-b2", "s-none"]);
    const result = await run(licenseReclaimAudit, only("b.example"));
    expect(keys(result)).toEqual(["s-b1", "s-b2"]);
    expect(tile(result, "Accounts")).toBe(2);
    expect(tile(result, "Licenses held")).toBe(2);
    expect(tile(result, "Disabled")).toBe(1);
    expect(result.notes[0]).toBe("Limited to accounts in: b.example.");
    // The product tally counts only the two accounts in scope.
    expect(result.notes).toContain(
      "By product: Microsoft 365 Business Basic 1, Microsoft 365 E3 1.",
    );
    expect(row(result.rows, "s-b1").cells.domain).toBe("b.example");
  });

  it("limits MFA registration, its tiles and its admin count", async () => {
    const everyone = await run(mfaRegistrationAudit, null);
    expect(keys(everyone)).toEqual(["s-a1", "s-a2", "s-b2", "s-c1", "s-none"]);
    expect(tile(everyone, "Admins without MFA")).toBe(2);
    const result = await run(mfaRegistrationAudit, only("a.example"));
    expect(keys(result)).toEqual(["s-a1", "s-a2"]);
    expect(tile(result, "Accounts")).toBe(2);
    expect(tile(result, "Not registered")).toBe(1);
    expect(tile(result, "Admins without MFA")).toBe(1);
    expect(result.notes[0]).toBe("Limited to accounts in: a.example.");
    expect(row(result.rows, "s-a2").cells.domain).toBe("a.example");
    expect(row(everyone.rows, "s-none").cells.domain).toBeNull();
    expect(keys(await run(mfaRegistrationAudit, only("c.example")))).toEqual(["s-c1"]);
  });

  it("leaves guests and license capacity as they are", async () => {
    for (const audit of [guestAccountsAudit, licenseCapacityAudit]) {
      const unscoped = await run(audit, null);
      const limited = await run(audit, only("a.example"));
      expect(limited, audit.definition.id).toEqual(unscoped);
      expect(limited.notes.join(" ")).not.toContain("Limited to");
    }
  });

  it("filters a table to one domain without changing the scope", async () => {
    const result = await run(inactiveAccountsAudit, null);
    const definition = auditDefinition(inactiveAccountsAudit);
    const filtered = queryAuditRows(definition, result.rows, {
      view: "all",
      facets: { domain: "b.example" },
    });
    expect(filtered.map((item) => item.key).sort()).toEqual(["s-b1", "s-b2"]);
  });

  it("keeps a very long list of domains inside one note", () => {
    const domains = new Set(Array.from({ length: 100 }, (_, index) => `domain-${index}.example`));
    const note = scopeNote(domains) ?? "";
    expect(note.length).toBeLessThanOrEqual(500);
    expect(note.startsWith("Limited to accounts in: domain-0.example, ")).toBe(true);
    expect(note).toMatch(/, and \d+ more\.$/);
    expect(scopeNote(new Set())).toBeNull();
    expect(scopeNote(null)).toBeNull();
  });

  it("counts each account once per domain, however many audits list it", () => {
    const account = (key: string, domain: unknown) => ({
      key,
      personId: null,
      tags: [],
      cells: { domain: domain as string },
    });
    const first = { rows: [account("1", "b.example"), account("2", "a.example")] };
    const second = {
      rows: [
        account("1", "b.example"),
        account("3", "b.example"),
        // Not listable: not a string, not a domain name, or missing.
        account("4", 12),
        account("5", "not a domain"),
        { key: "6", personId: null, tags: [], cells: {} },
      ],
    };
    expect(domainCounts([first, second])).toEqual([
      { domain: "a.example", accounts: 1 },
      { domain: "b.example", accounts: 2 },
    ]);
    expect(domainCounts([])).toEqual([]);
  });

  it("lists at most 1,000 domains, keeping the most populated", () => {
    const rows = Array.from({ length: 1_200 }, (_, index) => ({
      key: `k${index}`,
      personId: null,
      tags: [],
      cells: { domain: `d${String(index).padStart(4, "0")}.example` },
    }));
    // The first 100 domains hold a second account each, so they outrank the rest.
    const extra = rows.slice(0, 100).map((row) => ({ ...row, key: `${row.key}-again` }));
    const listed = domainCounts([{ rows: [...rows, ...extra] }]);
    expect(listed).toHaveLength(1_000);
    expect(listed.filter((item) => item.accounts === 2)).toHaveLength(100);
    expect(listed.map((item) => item.domain)).toEqual(listed.map((item) => item.domain).sort());
  });
});

describe("what the audit scope leaves out", () => {
  const registrations = [
    {
      id: "r-break-glass",
      userPrincipalName: "breakglass@tenant.onmicrosoft.example",
      userDisplayName: "Emergency access",
      isAdmin: true,
      isMfaRegistered: false,
      methodsRegistered: [],
      userType: "member",
    },
    {
      id: "r-ada",
      userPrincipalName: "ada@a.example",
      userDisplayName: "Ada Alpha",
      isAdmin: true,
      isMfaRegistered: true,
      methodsRegistered: ["microsoftAuthenticatorPush"],
      userType: "member",
    },
    {
      id: "r-bo",
      userPrincipalName: "bo@b.example",
      userDisplayName: "Bo Beta",
      isAdmin: false,
      isMfaRegistered: false,
      methodsRegistered: [],
      userType: "member",
    },
  ];
  const route: Route = (url) =>
    url.pathname === "/v1.0/reports/authenticationMethods/userRegistrationDetails"
      ? Response.json({ value: registrations })
      : tenant(url);
  const scoped = (scope: ReadonlySet<string> | null) =>
    mfaRegistrationAudit.run(context(mfaRegistrationAudit, route, undefined, scope));

  it("never shows a hidden admin without MFA as all clear", async () => {
    const result = await scoped(new Set(["a.example"]));
    const inScope = result.summary.find((item) => item.label === "Admins without MFA");
    expect(inScope).toMatchObject({ value: 0, tone: "neutral" });
    expect(
      result.summary.find((item) => item.label === "Admins without MFA outside the scope"),
    ).toMatchObject({
      value: 1,
      tone: "critical",
    });
    expect(result.notes.join(" ")).toContain("2 accounts in other domains were left out");
    expect(result.notes.join(" ")).toContain(
      "1 admin without MFA is in domains the audit scope leaves out",
    );
    // Without a scope the same tenant is critical in the main tile and has no outside tile.
    const all = await scoped(null);
    expect(all.summary.find((item) => item.label === "Admins without MFA")).toMatchObject({
      value: 1,
      tone: "critical",
    });
    expect(all.summary.some((item) => item.label.includes("outside the scope"))).toBe(false);
  });

  it("says what Licenses to reclaim left out, and reports every domain it read", async () => {
    const reclaim = await licenseReclaimAudit.run(
      context(licenseReclaimAudit, tenant, undefined, new Set(["nowhere.example"])),
    );
    expect(reclaim.rows).toEqual([]);
    expect(reclaim.notes.join(" ")).toContain(`${users.length} accounts in other domains`);
    expect(reclaim.domains?.map((item) => item.domain)).toContain("example.invalid");
  });

  it("never hides an account with attempts and no success behind the scope", async () => {
    const result = await inactiveAccountsAudit.run(
      context(inactiveAccountsAudit, tenant, undefined, new Set(["nowhere.example"])),
    );
    // u-failing tried to sign in three days ago with no success on record.
    expect(
      result.summary.find((item) => item.label === "Attempts, no success outside the scope"),
    ).toMatchObject({ value: 1, tone: "critical" });
    expect(result.notes.join(" ")).toContain("1 account with sign-in attempts and no success");
    const unscoped = await inactiveAccountsAudit.run(context(inactiveAccountsAudit, tenant));
    expect(unscoped.summary.some((item) => item.label.includes("outside the scope"))).toBe(false);
  });

  it("reports every domain it read, whatever the scope", async () => {
    const result = await scoped(new Set(["a.example"]));
    expect(result.domains).toEqual([
      { domain: "a.example", accounts: 1 },
      { domain: "b.example", accounts: 1 },
      { domain: "tenant.onmicrosoft.example", accounts: 1 },
    ]);
    const inactive = await inactiveAccountsAudit.run(
      context(inactiveAccountsAudit, tenant, undefined, new Set(["nowhere.example"])),
    );
    expect(inactive.rows).toEqual([]);
    expect(inactive.domains?.map((item) => item.domain)).toContain("example.invalid");
    expect(inactive.notes.join(" ")).toContain(`${users.length} accounts in other domains`);
  });
});

describe("what a run may read", () => {
  it("opens only enabled optional packs, and only the permissions enabled packs carry", () => {
    // LicenseAssignment.Read.All is granted for the enabled per-user licence pack, but the
    // inactive audit's optional tenant licence pack is off: the shared permission must not open it.
    const off = auditReadScope(
      inactiveAccountsAudit,
      ["directory.users.read", "audit.signins.read", "user.licenses.read"],
      ["User.Read.All", "AuditLog.Read.All", "LicenseAssignment.Read.All"],
    );
    expect(off.packs).not.toContain("tenant.licenses.read");
    expect(off.granted).not.toContain("LicenseAssignment.Read.All");
    expect(off.granted).toEqual(["User.Read.All", "AuditLog.Read.All"]);
    const on = auditReadScope(
      inactiveAccountsAudit,
      ["directory.users.read", "audit.signins.read", "tenant.licenses.read"],
      ["User.Read.All", "AuditLog.Read.All", "LicenseAssignment.Read.All"],
    );
    expect(on.packs).toContain("tenant.licenses.read");
    expect(on.granted).toContain("LicenseAssignment.Read.All");
  });
});
