import { auditRowSchema, auditSummaryItemSchema, queryAuditRows } from "@prism365/contracts";
import { GraphHttpClient, GraphRequestError } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  type Audit,
  type AuditContext,
  auditDefinition,
  builtInAudits,
  conditionalAccessAudit,
  createConditionalAccessAudit,
  type Pause,
  REQUEST_SPACING_MILLISECONDS,
} from "../src/index.ts";

// Synthetic tenant data only: invented policy names and ids, example.invalid domains.
const now = new Date("2026-10-22T12:00:00Z");
const GLOBAL_ADMIN = "62e90394-69f5-4237-9190-012177145e10";
const SECURITY_ADMIN = "194ae4cb-b126-40b2-bd5b-6091b380977d";
const UNLISTED_ROLE = "0e9cc0a1-0000-4000-8000-000000000042";
const BREAK_GLASS_ONE = "b1000000-0000-4000-8000-000000000001";
const BREAK_GLASS_TWO = "b1000000-0000-4000-8000-000000000002";
const GONE = "b1000000-0000-4000-8000-000000000003";
const GROUP = "9a000000-0000-4000-8000-000000000001";
const OFFICE = "10c00000-0000-4000-8000-000000000001";
const COUNTRIES = "10c00000-0000-4000-8000-000000000002";
const MISSING_LOCATION = "10c00000-0000-4000-8000-000000000003";
const SKU = "5c000000-0000-4000-8000-000000000001";

type Json = Record<string, unknown>;

const conditions = (
  over: {
    users?: Json;
    applications?: Json;
    clientAppTypes?: string[];
    locations?: Json | null;
    platforms?: Json | null;
    signInRiskLevels?: string[];
    userRiskLevels?: string[];
  } = {},
) => ({
  users: {
    includeUsers: ["All"],
    excludeUsers: [],
    includeGroups: [],
    excludeGroups: [],
    includeRoles: [],
    excludeRoles: [],
    ...over.users,
  },
  applications: { includeApplications: ["All"], excludeApplications: [], ...over.applications },
  clientAppTypes: over.clientAppTypes ?? ["all"],
  locations: over.locations ?? null,
  platforms: over.platforms ?? null,
  signInRiskLevels: over.signInRiskLevels ?? [],
  userRiskLevels: over.userRiskLevels ?? [],
});

const mfa = { operator: "OR", builtInControls: ["mfa"], authenticationStrength: null };
const block = { operator: "OR", builtInControls: ["block"], authenticationStrength: null };

const policy = (id: string, over: Json = {}) => ({
  id,
  displayName: `Policy ${id}`,
  state: "enabled",
  createdDateTime: "2026-01-05T10:00:00Z",
  modifiedDateTime: "2026-09-01T10:00:00Z",
  conditions: conditions(),
  grantControls: mfa,
  sessionControls: null,
  ...over,
});

// The two policies a healthy tenant has: legacy authentication blocked, MFA for everyone.
const legacyBlock = (over: Json = {}) =>
  policy("legacy", {
    conditions: conditions({ clientAppTypes: ["exchangeActiveSync", "other"] }),
    grantControls: block,
    ...over,
  });
const mfaForAll = (over: Json = {}) => policy("mfa-all", over);

const locations = [
  {
    "@odata.type": "#microsoft.graph.ipNamedLocation",
    id: OFFICE,
    displayName: "Head office",
    isTrusted: true,
    ipRanges: [{ "@odata.type": "#microsoft.graph.iPv4CidrRange", cidrAddress: "203.0.113.0/24" }],
  },
  {
    "@odata.type": "#microsoft.graph.countryNamedLocation",
    id: COUNTRIES,
    displayName: "Approved countries",
    countriesAndRegions: ["US", "CA"],
  },
];

const skus = (planName: string, capabilityStatus = "Enabled") => [
  {
    skuId: SKU,
    skuPartNumber: "TEST_SKU",
    capabilityStatus,
    servicePlans: [
      { servicePlanName: "EXCHANGE_S_STANDARD", provisioningStatus: "Success" },
      { servicePlanName: planName, provisioningStatus: "Success" },
    ],
  },
];

interface Tenant {
  policies?: unknown[];
  locations?: unknown[];
  skus?: unknown[];
  names?: Record<string, { displayName?: string; userPrincipalName?: string }>;
}

type Route = (url: URL) => Response | undefined;

function tenantRoute(tenant: Tenant): Route {
  return (url) => {
    if (url.pathname === "/v1.0/identity/conditionalAccess/policies")
      return Response.json({ value: tenant.policies ?? [] });
    if (url.pathname === "/v1.0/identity/conditionalAccess/namedLocations")
      return Response.json({ value: tenant.locations ?? locations });
    if (url.pathname === "/v1.0/subscribedSkus") return Response.json({ value: tenant.skus ?? [] });
    const user = /^\/v1\.0\/users\/([^/]+)$/.exec(url.pathname)?.[1];
    if (user) {
      const found = tenant.names?.[user];
      return found
        ? Response.json({ id: user, ...found })
        : Response.json({ error: { code: "Request_ResourceNotFound" } }, { status: 404 });
    }
    return undefined;
  };
}

const allGranted = ["Policy.Read.All", "User.Read.All", "LicenseAssignment.Read.All"];

function setup(
  route: Route,
  options: { granted?: string[]; scope?: ReadonlySet<string> | null; audit?: Audit } = {},
) {
  const trace: string[] = [];
  const pause: Pause = async (milliseconds) => {
    trace.push(`pause ${milliseconds}`);
  };
  const audit = options.audit ?? createConditionalAccessAudit({ pause });
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    // The client's own retry back-off is not under test here.
    sleep: async () => {},
    maxResponseBytes: 8_388_608,
    fetchImplementation: async (input) => {
      const url = new URL(String(input));
      trace.push(`GET ${url.pathname.replace("/v1.0/", "")}`);
      return route(url) ?? new Response("{}", { status: 404 });
    },
  });
  const context: AuditContext = {
    graph: scopeGraph(client, [...audit.packs, ...audit.optionalPacks]),
    now,
    granted: new Set(options.granted ?? allGranted),
    signal: new AbortController().signal,
    scope: options.scope ?? null,
  };
  return { audit, context, trace };
}

async function runWith(tenant: Tenant, options: Parameters<typeof setup>[1] = {}) {
  const { audit, context, trace } = setup(tenantRoute(tenant), options);
  return { result: await audit.run(context), trace };
}

const row = <T extends { key: string }>(rows: T[], key: string): T => {
  const found = rows.find((item) => item.key === key);
  if (!found) throw new Error(`No row ${key}`);
  return found;
};
const tile = (summary: { label: string; value: number }[], label: string) =>
  summary.find((item) => item.label === label);

const NO_LEGACY_BLOCK = "No enabled policy blocks legacy authentication";
const NO_ADMIN_MFA = "No enabled policy requires MFA for administrators";
const EXCLUDED_FROM_MFA = "Users excluded from MFA policies";

describe("the Conditional Access audit definition", () => {
  it("is in the catalog with its packs, permissions and no scope", () => {
    expect(builtInAudits).toContain(conditionalAccessAudit);
    const definition = auditDefinition(conditionalAccessAudit);
    expect(definition).toMatchObject({
      id: "identity.conditional-access",
      category: "Identity",
      scoped: false,
      permissions: ["Policy.Read.All"],
      optionalPermissions: ["User.Read.All", "LicenseAssignment.Read.All"],
    });
    expect(definition.views.map((view) => view.id)).toEqual([
      "all",
      "enabled",
      "report-only",
      "off",
      "legacy-block",
      "mfa",
      "exclusions",
    ]);
  });

  it("names the permission, roles and license Microsoft needs when it refuses", () => {
    expect(conditionalAccessAudit.forbidden).toContain("Policy.Read.All");
    expect(conditionalAccessAudit.forbidden).toContain("Security Reader");
    expect(conditionalAccessAudit.forbidden).toContain("Conditional Access Administrator");
    expect(conditionalAccessAudit.forbidden).toContain("Microsoft Entra ID P1");
  });

  it("paces real runs at more than one second between requests", () => {
    // The exported audit uses the real wait; tests above replace it.
    expect(REQUEST_SPACING_MILLISECONDS).toBeGreaterThan(1_000);
  });
});

describe("a tenant with legacy authentication blocked and MFA for everyone", () => {
  it("shows no warning tiles and valid rows", async () => {
    const { result } = await runWith({ policies: [legacyBlock(), mfaForAll()] });
    expect(result.summary.map((item) => auditSummaryItemSchema.parse(item).label)).toEqual([
      "Policies enabled",
      "Policies in report-only",
      "Policies off",
      EXCLUDED_FROM_MFA,
    ]);
    expect(tile(result.summary, "Policies enabled")?.value).toBe(2);
    expect(row(result.rows, "legacy").tags).toEqual(["enabled", "legacy-block"]);
    expect(row(result.rows, "mfa-all").tags).toEqual(["enabled", "mfa"]);
    for (const item of result.rows) auditRowSchema.parse(item);
    expect(result.complete).toBe(true);
  });

  it("counts policies by state and fills the state views", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        mfaForAll(),
        policy("watching", { state: "enabledForReportingButNotEnforced" }),
        policy("parked", { state: "disabled" }),
        policy("parked-too", { state: "disabled" }),
      ],
    });
    expect(tile(result.summary, "Policies enabled")?.value).toBe(2);
    expect(tile(result.summary, "Policies in report-only")?.value).toBe(1);
    expect(tile(result.summary, "Policies off")?.value).toBe(2);
    const definition = auditDefinition(conditionalAccessAudit);
    const states = (view: string) =>
      queryAuditRows(definition, result.rows, { view }).map((item) => item.cells.state);
    expect(states("enabled")).toEqual(["On", "On"]);
    expect(states("report-only")).toEqual(["Report only"]);
    expect(states("off")).toEqual(["Off", "Off"]);
  });

  it("lists a state Microsoft may add later without counting it as on, report-only or off", async () => {
    const { result } = await runWith({
      policies: [legacyBlock(), mfaForAll(), policy("odd", { state: "unknownFutureValue" })],
    });
    expect(row(result.rows, "odd").cells.state).toBe("Unknown");
    expect(row(result.rows, "odd").tags).not.toContain("enabled");
    expect(result.notes.join(" ")).toContain("1 policy has a state Prism365 does not recognise");
  });
});

describe("blocking legacy authentication", () => {
  it("is detected for an enabled policy that blocks both legacy client types for all users", async () => {
    const { result } = await runWith({ policies: [legacyBlock(), mfaForAll()] });
    expect(row(result.rows, "legacy").tags).toContain("legacy-block");
    expect(tile(result.summary, NO_LEGACY_BLOCK)).toBeUndefined();
  });

  it("is not detected when the only blocking policy is report-only", async () => {
    const { result } = await runWith({
      policies: [legacyBlock({ state: "enabledForReportingButNotEnforced" }), mfaForAll()],
    });
    expect(row(result.rows, "legacy").tags).not.toContain("legacy-block");
    expect(tile(result.summary, NO_LEGACY_BLOCK)).toMatchObject({
      value: 1,
      tone: "critical",
    });
  });

  it("is not detected when the only blocking policy is off", async () => {
    const { result } = await runWith({
      policies: [legacyBlock({ state: "disabled" }), mfaForAll()],
    });
    expect(tile(result.summary, NO_LEGACY_BLOCK)?.value).toBe(1);
  });

  it.each([
    ["only Exchange ActiveSync", ["exchangeActiveSync"]],
    ["only other clients", ["other"]],
    ["only browsers", ["browser"]],
    ["no client types", []],
  ])("is not detected for a partial policy: %s", async (_name, clientAppTypes) => {
    const { result } = await runWith({
      policies: [legacyBlock({ conditions: conditions({ clientAppTypes }) }), mfaForAll()],
    });
    expect(tile(result.summary, NO_LEGACY_BLOCK)?.value).toBe(1);
  });

  it("is not detected when the policy covers a group instead of all users", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock({
          conditions: conditions({
            users: { includeUsers: [], includeGroups: [GROUP] },
            clientAppTypes: ["exchangeActiveSync", "other"],
          }),
        }),
        mfaForAll(),
      ],
    });
    expect(tile(result.summary, NO_LEGACY_BLOCK)?.value).toBe(1);
  });

  it("is not detected when the policy grants access instead of blocking", async () => {
    const { result } = await runWith({
      policies: [legacyBlock({ grantControls: mfa }), mfaForAll()],
    });
    expect(tile(result.summary, NO_LEGACY_BLOCK)?.value).toBe(1);
  });

  it("does not count a block on all client apps, which another condition usually narrows", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock({
          conditions: conditions({
            clientAppTypes: ["all"],
            locations: { includeLocations: ["All"], excludeLocations: [OFFICE] },
          }),
        }),
        mfaForAll(),
      ],
    });
    expect(row(result.rows, "legacy").tags).not.toContain("legacy-block");
    expect(tile(result.summary, NO_LEGACY_BLOCK)?.value).toBe(1);
  });

  it("is detected when another policy blocks it and one is only report-only", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock({ id: "legacy-trial", state: "enabledForReportingButNotEnforced" }),
        legacyBlock(),
        mfaForAll(),
      ],
    });
    expect(tile(result.summary, NO_LEGACY_BLOCK)).toBeUndefined();
  });
});

describe("requiring MFA for administrators", () => {
  const adminsOnly = (over: Json = {}) =>
    policy("admins", {
      conditions: conditions({ users: { includeUsers: [], includeRoles: [GLOBAL_ADMIN] } }),
      ...over,
    });

  it("is met by an enabled MFA policy for all users", async () => {
    const { result } = await runWith({ policies: [legacyBlock(), mfaForAll()] });
    expect(tile(result.summary, NO_ADMIN_MFA)).toBeUndefined();
  });

  it("is met by an enabled MFA policy for the Global Administrator role", async () => {
    const { result } = await runWith({ policies: [legacyBlock(), adminsOnly()] });
    expect(tile(result.summary, NO_ADMIN_MFA)).toBeUndefined();
  });

  it("is met by an authentication strength in place of the MFA control", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        adminsOnly({
          grantControls: {
            operator: "AND",
            builtInControls: [],
            authenticationStrength: { id: "strength-1", displayName: "Phishing-resistant MFA" },
          },
        }),
      ],
    });
    expect(tile(result.summary, NO_ADMIN_MFA)).toBeUndefined();
    expect(row(result.rows, "admins").tags).toContain("mfa");
    expect(row(result.rows, "admins").cells.grant).toBe(
      "Require authentication strength Phishing-resistant MFA",
    );
  });

  it("is not met when the policy leaves the Global Administrator role out", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        mfaForAll({
          conditions: conditions({ users: { excludeRoles: [GLOBAL_ADMIN] } }),
        }),
      ],
    });
    expect(tile(result.summary, NO_ADMIN_MFA)).toMatchObject({ value: 1, tone: "warning" });
  });

  it("is not met when a policy for the Global Administrator role also excludes it", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        adminsOnly({
          conditions: conditions({
            users: { includeUsers: [], includeRoles: [GLOBAL_ADMIN], excludeRoles: [GLOBAL_ADMIN] },
          }),
        }),
      ],
    });
    expect(tile(result.summary, NO_ADMIN_MFA)?.value).toBe(1);
  });

  it("is met when another policy covers administrators and one excludes the role", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        mfaForAll({ conditions: conditions({ users: { excludeRoles: [GLOBAL_ADMIN] } }) }),
        adminsOnly(),
      ],
    });
    expect(tile(result.summary, NO_ADMIN_MFA)).toBeUndefined();
  });

  it("is not met by a report-only policy", async () => {
    const { result } = await runWith({
      policies: [legacyBlock(), mfaForAll({ state: "enabledForReportingButNotEnforced" })],
    });
    expect(tile(result.summary, NO_ADMIN_MFA)?.value).toBe(1);
  });

  it("is not met by a policy for a different administrator role", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        adminsOnly({
          conditions: conditions({ users: { includeUsers: [], includeRoles: [SECURITY_ADMIN] } }),
        }),
      ],
    });
    expect(tile(result.summary, NO_ADMIN_MFA)?.value).toBe(1);
  });

  it("is not met by a policy that asks for a compliant device but not MFA", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        mfaForAll({
          grantControls: { operator: "OR", builtInControls: ["compliantDevice"] },
        }),
      ],
    });
    expect(tile(result.summary, NO_ADMIN_MFA)?.value).toBe(1);
    expect(row(result.rows, "mfa-all").tags).not.toContain("mfa");
  });

  it("explains that security defaults are separate when a tile fires", async () => {
    const { result } = await runWith({ policies: [mfaForAll({ state: "disabled" })] });
    expect(tile(result.summary, NO_LEGACY_BLOCK)?.value).toBe(1);
    expect(result.notes.join(" ")).toContain("Security defaults are not Conditional Access");
  });
});

describe("exclusions", () => {
  const withExclusions = (id: string, users: Json, over: Json = {}) =>
    policy(id, { conditions: conditions({ users }), ...over });

  it("counts excluded users, groups and roles on each row and tags the policy", async () => {
    const { result } = await runWith(
      {
        policies: [
          legacyBlock(),
          withExclusions("mfa-all", {
            excludeUsers: [BREAK_GLASS_ONE, BREAK_GLASS_TWO],
            excludeGroups: [GROUP],
            excludeRoles: [GLOBAL_ADMIN, UNLISTED_ROLE],
          }),
        ],
      },
      { granted: ["Policy.Read.All"] },
    );
    const excluded = row(result.rows, "mfa-all");
    expect(excluded.tags).toContain("exclusions");
    expect(excluded.cells).toMatchObject({
      excludedUsers: 2,
      excludedGroups: 1,
      excludedRoles: 2,
    });
    expect(excluded.cells.excludes).toEqual([
      "2 users",
      "1 group",
      "2 roles (Global Administrator and 1 other)",
    ]);
    expect(row(result.rows, "legacy").tags).not.toContain("exclusions");
    expect(row(result.rows, "legacy").cells.excludes).toEqual([]);
  });

  it("counts distinct users excluded from enabled MFA policies only", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock({
          conditions: conditions({
            users: { excludeUsers: [GONE] },
            clientAppTypes: ["exchangeActiveSync", "other"],
          }),
        }),
        withExclusions("mfa-all", { excludeUsers: [BREAK_GLASS_ONE, BREAK_GLASS_TWO] }),
        withExclusions("mfa-admins", {
          includeUsers: [],
          includeRoles: [GLOBAL_ADMIN],
          excludeUsers: [BREAK_GLASS_ONE],
        }),
        withExclusions(
          "mfa-trial",
          { excludeUsers: ["b1000000-0000-4000-8000-0000000000aa"] },
          { state: "enabledForReportingButNotEnforced" },
        ),
      ],
    });
    // The block policy asks for no MFA and the trial is report-only, so neither counts.
    expect(tile(result.summary, EXCLUDED_FROM_MFA)?.value).toBe(2);
    expect(result.notes.join(" ")).toContain("break-glass");
  });

  it("does not count the literal values Microsoft uses in place of user ids", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        withExclusions("mfa-all", { excludeUsers: ["GuestsOrExternalUsers"] }),
      ],
    });
    expect(tile(result.summary, EXCLUDED_FROM_MFA)?.value).toBe(0);
    expect(row(result.rows, "mfa-all").cells).toMatchObject({
      excludedUsers: 0,
      excludes: ["Guests and external users"],
    });
    expect(row(result.rows, "mfa-all").tags).toContain("exclusions");
  });

  it("shows a policy that excludes nobody as leaving nobody out", async () => {
    const { result } = await runWith({
      policies: [legacyBlock(), withExclusions("mfa-all", { excludeUsers: ["None"] })],
    });
    expect(row(result.rows, "mfa-all").tags).not.toContain("exclusions");
  });

  it("fills the exclusions view", async () => {
    const { result } = await runWith({
      policies: [legacyBlock(), withExclusions("mfa-all", { excludeGroups: [GROUP] })],
    });
    const definition = auditDefinition(conditionalAccessAudit);
    expect(
      queryAuditRows(definition, result.rows, { view: "exclusions" }).map((item) => item.key),
    ).toEqual(["mfa-all"]);
  });
});

describe("who excluded users are", () => {
  const exclusions = [
    legacyBlock(),
    policy("mfa-all", {
      conditions: conditions({ users: { excludeUsers: [BREAK_GLASS_ONE, BREAK_GLASS_TWO, GONE] } }),
    }),
  ];
  const names = {
    [BREAK_GLASS_ONE]: { displayName: "Emergency Access One" },
    [BREAK_GLASS_TWO]: { userPrincipalName: "emergency2@example.invalid" },
  };

  it("names them when the session has granted User.Read.All", async () => {
    const { result, trace } = await runWith({ policies: exclusions, names });
    expect(row(result.rows, "mfa-all").cells.excludes).toEqual([
      "3 users: Emergency Access One, emergency2@example.invalid, 1 deleted or unknown account",
    ]);
    expect(trace.filter((entry) => entry.startsWith("GET users/")).sort()).toEqual([
      `GET users/${BREAK_GLASS_ONE}`,
      `GET users/${BREAK_GLASS_TWO}`,
      `GET users/${GONE}`,
    ]);
    expect(result.complete).toBe(true);
  });

  it("asks Microsoft only for the name fields", async () => {
    const requested: URL[] = [];
    const { audit, context } = setup((url) => {
      if (url.pathname.startsWith("/v1.0/users/")) requested.push(url);
      return tenantRoute({ policies: exclusions, names })(url);
    });
    await audit.run(context);
    expect(requested[0]?.searchParams.get("$select")).toBe("id,displayName,userPrincipalName");
  });

  it("shows counts and makes no user request without the optional pack", async () => {
    const { result, trace } = await runWith(
      { policies: exclusions, names },
      { granted: ["Policy.Read.All"] },
    );
    expect(row(result.rows, "mfa-all").cells.excludes).toEqual(["3 users"]);
    expect(trace.some((entry) => entry.startsWith("GET users/"))).toBe(false);
    expect(result.notes.join(" ")).toContain("User.Read.All");
    expect(JSON.stringify(result.rows)).not.toContain("Emergency Access One");
  });

  it("keeps the audit partial when a name lookup fails for another reason", async () => {
    const route = tenantRoute({ policies: exclusions, names });
    const { audit, context } = setup((url) =>
      url.pathname === `/v1.0/users/${BREAK_GLASS_TWO}`
        ? new Response("{}", { status: 500 })
        : route(url),
    );
    const result = await audit.run(context);
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("could not be named");
    expect(row(result.rows, "mfa-all").cells.excludes).toEqual([
      "3 users: Emergency Access One, 1 deleted or unknown account",
    ]);
  });

  it("looks up at most 50 users and says so", async () => {
    const many = Array.from(
      { length: 60 },
      (_, index) => `b2000000-0000-4000-8000-${String(index).padStart(12, "0")}`,
    );
    const { result, trace } = await runWith({
      policies: [
        legacyBlock(),
        policy("mfa-all", { conditions: conditions({ users: { excludeUsers: many } }) }),
      ],
      names: Object.fromEntries(many.map((id, index) => [id, { displayName: `User ${index}` }])),
    });
    expect(trace.filter((entry) => entry.startsWith("GET users/"))).toHaveLength(50);
    expect(result.notes.join(" ")).toContain("first 50 excluded users");
    expect(row(result.rows, "mfa-all").cells.excludedUsers).toBe(60);
  });

  it("never puts a value that is not a user id into a request path", async () => {
    const { trace } = await runWith({
      policies: [
        legacyBlock(),
        policy("mfa-all", {
          conditions: conditions({ users: { excludeUsers: ["../servicePrincipals", "x?y=z"] } }),
        }),
      ],
    });
    expect(trace.some((entry) => entry.startsWith("GET users/"))).toBe(false);
  });
});

describe("named locations", () => {
  const withLocations = policy("geo", {
    conditions: conditions({
      locations: { includeLocations: ["All"], excludeLocations: [OFFICE, COUNTRIES] },
    }),
  });

  it("resolves location ids to names", async () => {
    const { result } = await runWith({ policies: [legacyBlock(), mfaForAll(), withLocations] });
    expect(row(result.rows, "geo").cells.conditions).toEqual([
      "Locations: any location, except Head office and Approved countries",
    ]);
  });

  it("reads trusted locations by the words Microsoft uses", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        policy("trusted", {
          conditions: conditions({
            locations: { includeLocations: ["All"], excludeLocations: ["AllTrusted"] },
          }),
        }),
      ],
    });
    expect(row(result.rows, "trusted").cells.conditions).toEqual([
      "Locations: any location, except trusted locations",
    ]);
  });

  it("says when a location id is not among the named locations", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        policy("lost", {
          conditions: conditions({
            locations: { includeLocations: ["All"], excludeLocations: [MISSING_LOCATION] },
          }),
        }),
      ],
    });
    expect(row(result.rows, "lost").cells.conditions).toEqual([
      "Locations: any location, except 1 location that could not be named",
    ]);
  });

  it("does not read named locations when no policy uses one", async () => {
    const { trace } = await runWith({ policies: [legacyBlock(), mfaForAll()] });
    expect(trace.some((entry) => entry.includes("namedLocations"))).toBe(false);
  });

  it("lists locations without names and says so when they cannot be read", async () => {
    const route = tenantRoute({ policies: [withLocations] });
    const { audit, context } = setup((url) =>
      url.pathname.endsWith("/namedLocations") ? new Response("{}", { status: 403 }) : route(url),
    );
    const result = await audit.run(context);
    expect(row(result.rows, "geo").cells.conditions).toEqual([
      "Locations: any location, except 2 locations that could not be named",
    ]);
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("Named locations could not be read");
  });
});

describe("what each policy row says", () => {
  it("summarises who it applies to, the apps, conditions, grant and session", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        policy("admins", {
          displayName: "MFA for admin portals",
          modifiedDateTime: null,
          conditions: conditions({
            users: {
              includeUsers: [],
              includeRoles: [GLOBAL_ADMIN, SECURITY_ADMIN, UNLISTED_ROLE],
              includeGroups: [GROUP],
            },
            applications: { includeApplications: ["MicrosoftAdminPortals"] },
            clientAppTypes: ["browser", "mobileAppsAndDesktopClients"],
            platforms: { includePlatforms: ["windows", "macOS"], excludePlatforms: [] },
            signInRiskLevels: ["high", "medium"],
          }),
          grantControls: {
            operator: "AND",
            builtInControls: ["mfa", "compliantDevice"],
            authenticationStrength: null,
          },
          sessionControls: {
            signInFrequency: {
              isEnabled: true,
              type: "hours",
              value: 4,
              frequencyInterval: "timeBased",
            },
            persistentBrowser: { isEnabled: true, mode: "never" },
          },
        }),
        policy("guests", {
          conditions: conditions({
            users: {
              includeUsers: [],
              includeGuestsOrExternalUsers: {
                guestOrExternalUserTypes: "b2bCollaborationGuest,internalGuest",
              },
            },
            applications: { includeApplications: ["Office365"], excludeApplications: [GROUP] },
          }),
          grantControls: {
            operator: "OR",
            builtInControls: ["mfa", "compliantDevice"],
            authenticationStrength: null,
          },
        }),
      ],
    });
    expect(row(result.rows, "legacy").cells).toMatchObject({
      policy: "Policy legacy",
      state: "On",
      appliesTo: "All users",
      apps: "All apps",
      conditions: ["Client apps: Exchange ActiveSync, Other clients"],
      grant: "Block access",
      session: null,
      modified: "2026-09-01T10:00:00Z",
      created: "2026-01-05T10:00:00Z",
    });
    expect(row(result.rows, "admins").cells).toMatchObject({
      policy: "MFA for admin portals",
      appliesTo: "1 group, 3 roles (Global Administrator, Security Administrator and 1 other)",
      apps: "Microsoft admin portals",
      conditions: [
        "Client apps: Browser, Mobile apps and desktop clients",
        "Platforms: Windows, macOS",
        "Sign-in risk: High, Medium",
      ],
      grant: "Require MFA and a compliant device",
      session: "Sign in every 4 hours; Persistent browser: never",
      modified: null,
    });
    expect(row(result.rows, "guests").cells).toMatchObject({
      appliesTo: "Guests and external users (B2B collaboration guests and internal guests)",
      apps: "Office 365, except 1 app",
      grant: "Require one of: MFA, a compliant device",
    });
  });

  it("reads guest types sent as a list as well as a comma-separated string", async () => {
    const { result } = await runWith({
      policies: [
        policy("guests", {
          conditions: conditions({
            users: {
              includeUsers: [],
              includeGuestsOrExternalUsers: { guestOrExternalUserTypes: ["b2bDirectConnectUser"] },
            },
          }),
        }),
      ],
    });
    expect(row(result.rows, "guests").cells.appliesTo).toBe(
      "Guests and external users (B2B direct connect users)",
    );
  });

  it("sorts rows by policy name", async () => {
    const { result } = await runWith({
      policies: [
        policy("b", { displayName: "Beta" }),
        policy("a", { displayName: "Alpha" }),
        policy("c", { displayName: "Gamma" }),
      ],
    });
    expect(result.rows.map((item) => item.cells.policy)).toEqual(["Alpha", "Beta", "Gamma"]);
  });

  it("cleans provider text", async () => {
    const { result } = await runWith({
      policies: [policy("odd", { displayName: "Policy\u0000 with NUL" })],
    });
    expect(row(result.rows, "odd").cells.policy).toBe("Policy with NUL");
  });

  it("keeps every note inside the length a run stores", async () => {
    const { result } = await runWith({ policies: [mfaForAll({ state: "disabled" })], names: {} });
    for (const note of result.notes) expect(note.length).toBeLessThanOrEqual(500);
    for (const note of result.notes) {
      expect(note).not.toMatch(/[—·!]/);
    }
  });

  it("does not use the audit scope", async () => {
    const { result } = await runWith(
      { policies: [legacyBlock(), mfaForAll()] },
      { scope: new Set(["only.example.invalid"]) },
    );
    expect(result.rows).toHaveLength(2);
    expect(result.domains).toBeUndefined();
    expect(result.notes.join(" ")).not.toContain("Limited to");
  });
});

describe("policies Microsoft sent in a shape Prism365 cannot read", () => {
  it("leaves out an unreadable policy, counts it and marks the result partial", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        mfaForAll(),
        { displayName: "no id", state: "enabled" },
        policy("no-state", { state: undefined }),
        policy("garbled", { conditions: "not an object" }),
      ],
    });
    expect(result.rows.map((item) => item.key).sort()).toEqual(["legacy", "mfa-all"]);
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("3 policies could not be read and are left out");
  });

  it("tolerates nulls and unknown fields in a readable policy", async () => {
    const { result } = await runWith({
      policies: [
        legacyBlock(),
        mfaForAll(),
        {
          id: "sparse",
          displayName: null,
          state: "enabled",
          createdDateTime: "not a date",
          modifiedDateTime: null,
          templateId: null,
          conditions: { users: { includeUsers: ["All"], excludeUsers: null }, futureField: 1 },
          grantControls: null,
          sessionControls: { disableResilienceDefaults: null, somethingNew: {} },
          somethingNew: true,
        },
      ],
    });
    expect(row(result.rows, "sparse").cells).toMatchObject({
      policy: "Unnamed policy",
      appliesTo: "All users",
      grant: "No grant control",
      created: null,
      modified: null,
    });
    expect(result.complete).toBe(true);
  });

  it("does not let one unreadable policy hide that nothing blocks legacy authentication", async () => {
    const { result } = await runWith({
      policies: [mfaForAll(), { displayName: "no id" }],
    });
    expect(tile(result.summary, NO_LEGACY_BLOCK)?.value).toBe(1);
    expect(result.notes.join(" ")).toContain("a check above may be missing a policy");
  });

  it("reads a list of one policy that Microsoft returns over several pages", async () => {
    const route = tenantRoute({ policies: [] });
    const { audit, context, trace } = setup((url) => {
      if (url.pathname === "/v1.0/identity/conditionalAccess/policies") {
        const second = url.searchParams.get("$skiptoken") === "2";
        return Response.json(
          second
            ? { value: [mfaForAll()] }
            : {
                value: [legacyBlock()],
                "@odata.nextLink":
                  "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$skiptoken=2",
              },
        );
      }
      return route(url);
    });
    const result = await audit.run(context);
    expect(result.rows).toHaveLength(2);
    expect(trace).toEqual([
      "GET identity/conditionalAccess/policies",
      `pause ${REQUEST_SPACING_MILLISECONDS}`,
      "GET identity/conditionalAccess/policies",
    ]);
  });
});

describe("when Microsoft refuses", () => {
  it("passes a 403 on the policy list up so it can be explained", async () => {
    const { audit, context } = setup(() => new Response("{}", { status: 403 }));
    await expect(audit.run(context)).rejects.toMatchObject({
      constructor: GraphRequestError,
      status: 403,
    });
  });

  it("passes a 429 that survived the client's retries up as throttling", async () => {
    const { audit, context } = setup(() => new Response("{}", { status: 429 }));
    await expect(audit.run(context)).rejects.toMatchObject({ status: 429 });
  });
});

describe("an empty policy list", () => {
  it("says the tenant has Entra ID P1 and shows the missing-policy tiles", async () => {
    const { result } = await runWith({ policies: [], skus: skus("AAD_PREMIUM") });
    expect(result.rows).toEqual([]);
    expect(result.notes[0]).toContain("has Microsoft Entra ID P1 or P2");
    expect(tile(result.summary, NO_LEGACY_BLOCK)?.value).toBe(1);
    expect(tile(result.summary, NO_ADMIN_MFA)?.value).toBe(1);
    expect(result.complete).toBe(true);
  });

  it("accepts the P2 plan as proof of P1", async () => {
    const { result } = await runWith({ policies: [], skus: skus("AAD_PREMIUM_P2") });
    expect(result.notes[0]).toContain("has Microsoft Entra ID P1 or P2");
  });

  it("says no subscription includes P1 and does not call a missing policy a finding", async () => {
    const { result } = await runWith({ policies: [], skus: skus("EXCHANGE_S_ENTERPRISE") });
    expect(result.notes[0]).toContain(
      "none of the tenant's subscriptions include Microsoft Entra ID P1",
    );
    expect(tile(result.summary, NO_LEGACY_BLOCK)).toBeUndefined();
    expect(tile(result.summary, NO_ADMIN_MFA)).toBeUndefined();
  });

  it("does not count a subscription that is suspended", async () => {
    const { result } = await runWith({
      policies: [],
      skus: skus("AAD_PREMIUM", "Suspended"),
    });
    expect(result.notes[0]).toContain("none of the tenant's subscriptions include");
  });

  it("says it could not check when the license pack is not granted", async () => {
    const { result, trace } = await runWith(
      { policies: [], skus: skus("AAD_PREMIUM") },
      { granted: ["Policy.Read.All"] },
    );
    expect(result.notes[0]).toContain("could not be checked");
    expect(result.notes[0]).toContain("LicenseAssignment.Read.All");
    expect(trace.some((entry) => entry.includes("subscribedSkus"))).toBe(false);
    expect(tile(result.summary, NO_LEGACY_BLOCK)).toBeUndefined();
  });

  it("says it could not check when the subscriptions cannot be read", async () => {
    const route = tenantRoute({ policies: [] });
    const { audit, context } = setup((url) =>
      url.pathname === "/v1.0/subscribedSkus" ? new Response("{}", { status: 403 }) : route(url),
    );
    const result = await audit.run(context);
    expect(result.notes[0]).toContain("could not be checked");
  });

  it("does not read subscriptions when policies came back", async () => {
    const { trace } = await runWith({ policies: [mfaForAll()], skus: skus("AAD_PREMIUM") });
    expect(trace.some((entry) => entry.includes("subscribedSkus"))).toBe(false);
  });
});

describe("pacing against Microsoft's one request per second limit", () => {
  it("waits between the policy list and the named locations, and not before the first request", async () => {
    const { trace } = await runWith({
      policies: [
        policy("geo", {
          conditions: conditions({ locations: { includeLocations: [OFFICE] } }),
        }),
      ],
    });
    expect(trace).toEqual([
      "GET identity/conditionalAccess/policies",
      `pause ${REQUEST_SPACING_MILLISECONDS}`,
      "GET identity/conditionalAccess/namedLocations",
    ]);
  });

  it("makes a single request with no wait when there is nothing more to read", async () => {
    const { trace } = await runWith({ policies: [mfaForAll()] });
    expect(trace).toEqual(["GET identity/conditionalAccess/policies"]);
  });

  it("does not wait for the endpoints that have their own limits", async () => {
    const { trace } = await runWith({
      policies: [
        policy("mfa-all", {
          conditions: conditions({ users: { excludeUsers: [BREAK_GLASS_ONE] } }),
        }),
      ],
      names: { [BREAK_GLASS_ONE]: { displayName: "Emergency Access One" } },
    });
    expect(trace).toEqual([
      "GET identity/conditionalAccess/policies",
      `GET users/${BREAK_GLASS_ONE}`,
    ]);
  });

  it("stops waiting when the audit is stopped", async () => {
    const controller = new AbortController();
    const audit = createConditionalAccessAudit({
      pause: async (_milliseconds, signal) => {
        controller.abort(new Error("time limit"));
        signal.throwIfAborted();
      },
    });
    const { context } = setup(
      tenantRoute({
        policies: [
          policy("geo", { conditions: conditions({ locations: { includeLocations: [OFFICE] } }) }),
        ],
      }),
      { audit },
    );
    await expect(audit.run({ ...context, signal: controller.signal })).rejects.toThrow(
      "time limit",
    );
  });
});
