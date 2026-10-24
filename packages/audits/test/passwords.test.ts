import { auditRowSchema, queryAuditRows } from "@prism365/contracts";
import { GraphHttpClient } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import { type AuditContext, auditDefinition, passwordsAudit } from "../src/index.ts";

// Synthetic tenant data only: invented names, example.invalid domains.
const now = new Date("2026-10-22T12:00:00Z");
const DAY = 86_400_000;
const daysAgo = (days: number) => new Date(now.getTime() - days * DAY).toISOString();
const plus = (iso: string, days: number) => new Date(Date.parse(iso) + days * DAY).toISOString();

type Json = Record<string, unknown>;

// One domain that leaves both settings unset, one that never expires, one with its own lifetime
// and window, and one whose lifetime is malformed.
const domains = [
  { id: "default.example.invalid" },
  { id: "never.example.invalid", passwordValidityPeriodInDays: 2_147_483_647 },
  {
    id: "long.example.invalid",
    passwordValidityPeriodInDays: 180,
    passwordNotificationWindowInDays: 30,
  },
  { id: "broken.example.invalid", passwordValidityPeriodInDays: "ninety" },
];

let counter = 0;
function member(domain: string, changedDaysAgo: number | null, over: Json = {}): Json {
  counter += 1;
  return {
    id: `u-${counter}`,
    displayName: `Member ${counter}`,
    userPrincipalName: `member${counter}@${domain}`,
    userType: "Member",
    accountEnabled: true,
    createdDateTime: daysAgo(900),
    lastPasswordChangeDateTime: changedDaysAgo === null ? null : daysAgo(changedDaysAgo),
    ...over,
  };
}

type Route = (url: URL) => Response | undefined;

function setup(
  users: Json[],
  options: {
    granted?: string[];
    scope?: ReadonlySet<string> | null;
    domainRoute?: Route;
    usersRoute?: Route;
  } = {},
) {
  const requests: URL[] = [];
  const route: Route = (url) => {
    if (url.pathname === "/v1.0/users") {
      const replaced = options.usersRoute?.(url);
      if (replaced) return replaced;
      // As Graph does: only the fields that were selected come back.
      const fields = url.searchParams.get("$select")?.split(",");
      return Response.json({
        value: fields
          ? users.map((user) =>
              Object.fromEntries(Object.entries(user).filter(([key]) => fields.includes(key))),
            )
          : users,
      });
    }
    if (url.pathname === "/v1.0/domains")
      return options.domainRoute?.(url) ?? Response.json({ value: domains });
    return undefined;
  };
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    maxResponseBytes: 8_388_608,
    maxRetries: 0,
    fetchImplementation: async (input) => {
      const url = new URL(String(input));
      requests.push(url);
      return route(url) ?? new Response("{}", { status: 404 });
    },
  });
  const context: AuditContext = {
    graph: scopeGraph(client, [...passwordsAudit.packs, ...passwordsAudit.optionalPacks]),
    now,
    granted: new Set(options.granted ?? ["User.Read.All", "Domain.Read.All"]),
    signal: new AbortController().signal,
    scope: options.scope ?? null,
  };
  return { context, requests };
}

async function run(users: Json[], options: Parameters<typeof setup>[1] = {}) {
  const { context, requests } = setup(users, options);
  const result = await passwordsAudit.run(context);
  const byId = (id: unknown) => {
    const row = result.rows.find((item) => item.key === id);
    if (!row) throw new Error(`No row ${String(id)}`);
    return row;
  };
  return { result, requests, byId };
}

const tile = (result: { summary: { label: string; value: number }[] }, label: string) =>
  result.summary.find((item) => item.label === label)?.value;

describe("the password status", () => {
  it("names the first of its rules that matches", async () => {
    const weakAndOld = member("default.example.invalid", 400, {
      passwordPolicies: "DisableStrongPassword",
    });
    const expired = member("default.example.invalid", 95);
    const expiring = member("default.example.invalid", 80);
    const unreadable = member("default.example.invalid", 5, {
      lastPasswordChangeDateTime: "last tuesday",
    });
    const never = member("default.example.invalid", null);
    const oldAndNeverExpires = member("never.example.invalid", 500);
    const fresh = member("default.example.invalid", 10);
    const { byId } = await run([
      weakAndOld,
      expired,
      expiring,
      unreadable,
      never,
      oldAndNeverExpires,
      fresh,
    ]);
    // A weak setting outranks an old, expired password.
    expect(byId(weakAndOld.id).cells.status).toBe("Strong password not required");
    expect(byId(expired.id).cells.status).toBe("Expired");
    expect(byId(expiring.id).cells.status).toBe("Expires soon");
    expect(byId(unreadable.id).cells.status).toBe("Unreadable date");
    expect(byId(never.id).cells.status).toBe("Never changed");
    expect(byId(oldAndNeverExpires.id).cells.status).toBe("Older than 1 year");
    expect(byId(fresh.id).cells.status).toBe("OK");
  });

  it("ranks expired above expiring soon above never changed above old", async () => {
    // Expired and old: expiry wins. Expiring and old: expiry wins. Never changed is not old.
    const expiredAndOld = member("default.example.invalid", 600);
    const expiringAndOld = member("long.example.invalid", 160);
    const { byId } = await run([expiredAndOld, expiringAndOld]);
    expect(byId(expiredAndOld.id).cells.status).toBe("Expired");
    expect(byId(expiringAndOld.id).cells.status).toBe("Expires soon");
  });

  it("names an old password by the largest age it passes", async () => {
    const users = [89, 90, 179, 180, 364, 365].map((age) => member("never.example.invalid", age));
    const { byId } = await run(users);
    expect(users.map((user) => byId(user.id).cells.status)).toEqual([
      "OK",
      "Older than 90 days",
      "Older than 90 days",
      "Older than 180 days",
      "Older than 180 days",
      "Older than 1 year",
    ]);
  });

  it("counts the age in whole days from the last change", async () => {
    const user = member("never.example.invalid", 100);
    const { byId } = await run([user]);
    expect(byId(user.id).cells).toMatchObject({ lastChanged: daysAgo(100), days: 100 });
  });
});

describe("the domain policy", () => {
  it("gives a domain that sets nothing Microsoft's 90 days and 14-day window", async () => {
    const lastChanged = daysAgo(80);
    const user = member("default.example.invalid", 80);
    const { byId } = await run([user]);
    expect(byId(user.id).cells).toMatchObject({
      policy: "90 days",
      expires: plus(lastChanged, 90),
      status: "Expires soon",
    });
  });

  it("starts the notification window exactly 14 days before the end, and not a day sooner", async () => {
    const edge = member("default.example.invalid", 76);
    const before = member("default.example.invalid", 75);
    const lastDay = member("default.example.invalid", 90);
    const { byId } = await run([edge, before, lastDay]);
    expect(byId(edge.id).cells.status).toBe("Expires soon");
    expect(byId(before.id).cells.status).toBe("OK");
    // A password that ends today has ended.
    expect(byId(lastDay.id).cells.status).toBe("Expired");
  });

  it("uses a domain's own lifetime and window", async () => {
    const soon = member("long.example.invalid", 150);
    const later = member("long.example.invalid", 149);
    const { byId } = await run([soon, later]);
    expect(byId(soon.id).cells).toMatchObject({ policy: "180 days", status: "Expires soon" });
    expect(byId(later.id).cells.status).toBe("Older than 90 days");
    expect(byId(later.id).tags).not.toContain("expiring");
  });

  it("reads 2147483647 days as never expires and gives no expiry date", async () => {
    const user = member("never.example.invalid", 700);
    const { byId, result } = await run([user]);
    expect(byId(user.id).cells).toMatchObject({ policy: "Never", expires: null });
    expect(tile(result, "Expired")).toBe(0);
  });

  it("matches the domain after the @ without regard to case", async () => {
    const user = member("default.example.invalid", 95, {
      userPrincipalName: "Mixed.Case@DEFAULT.Example.Invalid",
    });
    const { byId } = await run([user]);
    expect(byId(user.id).cells.status).toBe("Expired");
  });

  it("matches the domain Microsoft lists in any letter case", async () => {
    const user = member("upper.example.invalid", 40);
    const { byId } = await run([user], {
      domainRoute: () =>
        Response.json({
          value: [{ id: "Upper.Example.Invalid", passwordValidityPeriodInDays: 30 }],
        }),
    });
    expect(byId(user.id).cells).toMatchObject({ policy: "30 days", status: "Expired" });
  });

  it("leaves the policy unknown for a domain Microsoft did not list or sent unreadably", async () => {
    const stranger = member("elsewhere.example.invalid", 500);
    const broken = member("broken.example.invalid", 500);
    const { byId, result } = await run([stranger, broken]);
    for (const user of [stranger, broken])
      expect(byId(user.id).cells).toMatchObject({
        policy: "Unknown",
        expires: null,
        status: "Older than 1 year",
      });
    // One of four domain records was unreadable.
    expect(result.complete).toBe(false);
  });

  it("shows an account synced from on-premises as managed there, with no expiry date", async () => {
    const synced = member("default.example.invalid", 500, { onPremisesSyncEnabled: true });
    const cloud = member("default.example.invalid", 500);
    const { byId } = await run([synced, cloud]);
    expect(byId(synced.id).cells).toMatchObject({
      source: "On-premises",
      policy: "Managed on-premises",
      expires: null,
      status: "Older than 1 year",
    });
    expect(byId(synced.id).tags).not.toContain("expired");
    expect(byId(cloud.id).cells).toMatchObject({ source: "Cloud", status: "Expired" });
  });

  it("gives an account set never to expire no expiry date, whatever its domain says", async () => {
    const user = member("default.example.invalid", 500, {
      passwordPolicies: "DisablePasswordExpiration",
    });
    const { byId } = await run([user]);
    expect(byId(user.id).cells).toMatchObject({
      neverExpires: true,
      weak: false,
      policy: "90 days",
      expires: null,
      status: "Older than 1 year",
    });
  });

  it("reads both flags from one comma-separated value", async () => {
    const user = member("default.example.invalid", 10, {
      passwordPolicies: "DisablePasswordExpiration, DisableStrongPassword",
    });
    const { byId, result } = await run([user]);
    expect(byId(user.id).cells).toMatchObject({ neverExpires: true, weak: true });
    expect(tile(result, "Never expires")).toBe(1);
    expect(tile(result, "Strong password not required")).toBe(1);
  });

  it("does not take a flag from a longer word", async () => {
    const user = member("default.example.invalid", 10, {
      passwordPolicies: "NotDisableStrongPassword",
    });
    const { byId } = await run([user]);
    expect(byId(user.id).cells).toMatchObject({ weak: false, neverExpires: false });
  });

  it("does not turn a lifetime beyond the calendar into a date", async () => {
    const huge = {
      id: "huge.example.invalid",
      passwordValidityPeriodInDays: 2_147_483_000,
    };
    const user = member("huge.example.invalid", 10);
    const { byId, result } = await run([user], {
      domainRoute: () => Response.json({ value: [huge] }),
    });
    expect(byId(user.id).cells).toMatchObject({ policy: "2147483000 days", expires: null });
    expect(result.complete).toBe(true);
  });
});

describe("missing and damaged dates", () => {
  it("counts only enabled accounts in the never-changed tile", async () => {
    const enabled = member("default.example.invalid", null);
    const disabled = member("default.example.invalid", null, { accountEnabled: false });
    const { byId, result } = await run([enabled, disabled]);
    expect(byId(enabled.id).cells.status).toBe("Never changed");
    expect(byId(disabled.id).cells.status).toBe("Never changed");
    expect(tile(result, "Enabled, never changed")).toBe(1);
  });

  it("reads a missing field as never changed, but a damaged one as unreadable", async () => {
    const missing = member("default.example.invalid", null);
    delete missing.lastPasswordChangeDateTime;
    const damaged = member("default.example.invalid", 5, { lastPasswordChangeDateTime: 12 });
    const future = member("default.example.invalid", -5);
    const ancient = member("default.example.invalid", 40_000);
    const { byId, result } = await run([missing, damaged, future, ancient]);
    expect(byId(missing.id).cells.status).toBe("Never changed");
    for (const user of [damaged, future, ancient]) {
      expect(byId(user.id).cells).toMatchObject({
        status: "Unreadable date",
        lastChanged: null,
        days: null,
        expires: null,
      });
      expect(byId(user.id).tags).toEqual(expect.arrayContaining(["unreadable"]));
      expect(byId(user.id).tags).not.toContain("never-changed");
    }
    expect(result.complete).toBe(false);
    expect(result.notes.join("\n")).toContain("3 accounts have a password date");
  });

  it("allows a clock a few hours ahead of Microsoft's", async () => {
    const user = member("default.example.invalid", -0.5);
    const { byId, result } = await run([user], { granted: ["User.Read.All"] });
    expect(byId(user.id).cells).toMatchObject({ days: 0, status: "OK" });
    expect(result.complete).toBe(true);
  });
});

describe("what is listed", () => {
  it("lists members only and counts guests in a note", async () => {
    const people = [member("default.example.invalid", 10)];
    const guests = [
      member("tenant.example.invalid", 900, { userType: "Guest" }),
      member("tenant.example.invalid", 900, { userType: "guest" }),
    ];
    const { result } = await run([...people, ...guests]);
    expect(result.rows).toHaveLength(1);
    expect(tile(result, "Member accounts")).toBe(1);
    expect(result.notes.join("\n")).toContain("2 guest accounts are not listed");
  });

  it("keeps an account of an unknown type", async () => {
    const odd = member("default.example.invalid", 10, { userType: null });
    const { byId } = await run([odd]);
    expect(byId(odd.id).personId).toBe(odd.id);
  });

  it("limits rows, tiles and notes to the audit scope and names what it left out", async () => {
    const inside = member("never.example.invalid", 500, {
      passwordPolicies: "DisableStrongPassword",
    });
    const outside = member("default.example.invalid", 500, {
      passwordPolicies: "DisableStrongPassword",
    });
    const { result } = await run([inside, outside], {
      scope: new Set(["never.example.invalid"]),
    });
    expect(result.rows.map((row) => row.key)).toEqual([inside.id]);
    expect(tile(result, "Strong password not required")).toBe(1);
    expect(result.notes.join("\n")).toContain("Limited to accounts in: never.example.invalid.");
    expect(result.notes.join("\n")).toContain("1 account in another domain was left out");
    // The inventory covers every domain the run read, whatever the scope.
    expect(result.domains?.map((item) => item.domain).sort()).toEqual([
      "default.example.invalid",
      "never.example.invalid",
    ]);
  });
});

describe("what the audit asks Microsoft", () => {
  it("selects the password fields, asks for no sign-in dates and sends no $top to domains", async () => {
    const { requests } = await run([member("default.example.invalid", 10)], {
      granted: ["User.Read.All", "AuditLog.Read.All", "Domain.Read.All"],
    });
    const users = requests.find((url) => url.pathname === "/v1.0/users");
    const fields = users?.searchParams.get("$select")?.split(",") ?? [];
    expect(fields).toEqual(
      expect.arrayContaining([
        "lastPasswordChangeDateTime",
        "passwordPolicies",
        "accountEnabled",
        "userType",
        "onPremisesSyncEnabled",
        "createdDateTime",
        "userPrincipalName",
        "displayName",
        "id",
      ]),
    );
    expect(fields).not.toContain("signInActivity");
    for (const url of requests.filter((item) => item.pathname === "/v1.0/domains"))
      expect(url.searchParams.has("$top")).toBe(false);
  });

  it("works from password age alone when domains are not granted", async () => {
    const user = member("default.example.invalid", 500);
    const { result, requests, byId } = await run([user], { granted: ["User.Read.All"] });
    expect(requests.some((url) => url.pathname === "/v1.0/domains")).toBe(false);
    expect(byId(user.id).cells).toMatchObject({
      policy: "Unknown",
      expires: null,
      status: "Older than 1 year",
    });
    expect(result.summary.map((item) => item.label)).not.toContain("Expires soon");
    expect(result.complete).toBe(true);
    expect(result.notes.join("\n")).toContain("domains were not read");
  });

  it("goes on without expiry dates when Microsoft refuses the domain list, and says so", async () => {
    const user = member("default.example.invalid", 500);
    const { result, byId } = await run([user], {
      domainRoute: () => new Response("{}", { status: 403 }),
    });
    expect(byId(user.id).cells).toMatchObject({ policy: "Unknown", expires: null });
    expect(result.complete).toBe(false);
    expect(result.summary.map((item) => item.label)).not.toContain("Expires soon");
    expect(result.notes.join("\n")).toContain("refused to list domains");
  });

  it("falls back to the plain domain list when $select is rejected", async () => {
    const user = member("default.example.invalid", 80);
    const { result, requests, byId } = await run([user], {
      domainRoute: (url) =>
        url.searchParams.has("$select")
          ? new Response("{}", { status: 400 })
          : Response.json({ value: domains.slice(0, 3) }),
    });
    expect(requests.filter((url) => url.pathname === "/v1.0/domains")).toHaveLength(2);
    expect(byId(user.id).cells.status).toBe("Expires soon");
    expect(result.complete).toBe(true);
  });

  it("fails the run when Microsoft refuses the account list", async () => {
    const { context } = setup([], { usersRoute: () => new Response("{}", { status: 403 }) });
    await expect(passwordsAudit.run(context)).rejects.toMatchObject({ status: 403 });
  });
});

describe("the published table", () => {
  const users = [
    member("default.example.invalid", 500, { passwordPolicies: "DisableStrongPassword" }),
    member("default.example.invalid", 500, { passwordPolicies: "DisablePasswordExpiration" }),
    member("default.example.invalid", 95),
    member("default.example.invalid", 80),
    member("never.example.invalid", 120),
    member("never.example.invalid", 200),
    member("never.example.invalid", 400),
    member("never.example.invalid", null),
    // A disabled account that never changed its password is listed but not counted in the tile.
    member("never.example.invalid", null, { accountEnabled: false }),
    member("never.example.invalid", 20),
  ];

  it("has views, a sort and tiles that agree with each other and with the rows", async () => {
    const definition = auditDefinition(passwordsAudit);
    const { result } = await run(users);
    const keys = new Set(definition.columns.map((column) => column.key));
    expect(keys.has(definition.defaultSort.column)).toBe(true);
    expect(definition.defaultView).toBe("d365");
    expect(definition.scoped).toBe(true);
    expect(definition.handoff).toBe(true);
    expect(definition.permissions).toEqual(["User.Read.All"]);
    expect(definition.optionalPermissions).toEqual(["Domain.Read.All"]);
    for (const row of result.rows) {
      auditRowSchema.parse(row);
      for (const key of Object.keys(row.cells)) expect(keys.has(key), key).toBe(true);
    }
    const views = new Set(definition.views.map((view) => view.id));
    for (const item of result.summary)
      if (item.view) expect(views.has(item.view), item.label).toBe(true);
    // Each tile is the size of the view it opens.
    for (const item of result.summary) {
      if (!item.view) continue;
      const shown = queryAuditRows(definition, result.rows, { view: item.view });
      expect(shown.length, item.label).toBe(item.value);
    }
  });

  it("offers the three ages, each including every older password", async () => {
    const definition = auditDefinition(passwordsAudit);
    const { result } = await run(users);
    const sizes = Object.fromEntries(
      ["d90", "d180", "d365"].map((view) => [
        view,
        queryAuditRows(definition, result.rows, { view }).length,
      ]),
    );
    // 500, 500, 95, 120, 200, 400 are 90+; 500, 500, 200, 400 are 180+; 500, 500, 400 are 365+.
    expect(sizes).toEqual({ d90: 6, d180: 4, d365: 3 });
  });

  it("counts each tile", async () => {
    const { result } = await run(users);
    expect(Object.fromEntries(result.summary.map((item) => [item.label, item.value]))).toEqual({
      "Member accounts": 10,
      "Strong password not required": 1,
      "Never expires": 1,
      "Older than 1 year": 3,
      Expired: 2,
      "Expires soon": 1,
      "Enabled, never changed": 1,
    });
  });

  it("keeps every note under the 500 characters the result stores", async () => {
    const { result } = await run(users, { granted: ["User.Read.All"] });
    for (const note of result.notes) expect(note.length, note).toBeLessThanOrEqual(500);
    const { result: withDomains } = await run(users);
    for (const note of withDomains.notes) expect(note.length, note).toBeLessThanOrEqual(500);
  });
});
