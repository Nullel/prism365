import { auditRowSchema, auditSummaryItemSchema, queryAuditRows } from "@prism365/contracts";
import { GraphHttpClient, GraphRequestError } from "@prism365/graph";
import { scopeGraph } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  type Audit,
  type AuditContext,
  auditDefinition,
  builtInAudits,
  createRiskyUsersAudit,
  type Pause,
  REQUEST_SPACING_MILLISECONDS,
  riskyUsersAudit,
} from "../src/index.ts";

// Synthetic tenant data only: invented names, example.invalid domains.
const now = new Date("2026-10-22T12:00:00Z");
const daysAgo = (days: number) => new Date(now.getTime() - days * 86_400_000).toISOString();
const P2 = "Risky user data in Microsoft Graph needs Microsoft Entra ID P2.";

type Json = Record<string, unknown>;

const risky = (id: string, over: Json = {}) => ({
  id,
  isDeleted: false,
  isProcessing: false,
  riskLastUpdatedDateTime: daysAgo(2),
  riskLevel: "high",
  riskState: "atRisk",
  riskDetail: "none",
  userDisplayName: `User ${id}`,
  userPrincipalName: `${id}@example.invalid`,
  ...over,
});

const tenant = [
  risky("high-open"),
  risky("medium-open", { riskLevel: "medium", riskLastUpdatedDateTime: daysAgo(5) }),
  risky("low-open", { riskLevel: "low" }),
  risky("compromised", {
    riskLevel: "high",
    riskState: "confirmedCompromised",
    riskDetail: "adminConfirmedUserCompromised",
    isProcessing: true,
  }),
  risky("fixed-recently", {
    riskLevel: "none",
    riskState: "remediated",
    riskDetail: "userPerformedSecuredPasswordReset",
    riskLastUpdatedDateTime: daysAgo(9),
  }),
  risky("dismissed-old", {
    riskLevel: "none",
    riskState: "dismissed",
    riskDetail: "adminDismissedAllRiskForUser",
    riskLastUpdatedDateTime: daysAgo(80),
  }),
  risky("safe", {
    riskLevel: "none",
    riskState: "confirmedSafe",
    riskDetail: "adminConfirmedSigninSafe",
  }),
  risky("uncovered", { riskLevel: "hidden", riskState: "none", riskDetail: "hidden" }),
];

type Route = (url: URL) => Response | undefined;

function setup(route: Route, options: { scope?: ReadonlySet<string> | null; audit?: Audit } = {}) {
  const trace: string[] = [];
  const urls: URL[] = [];
  const pause: Pause = async (milliseconds) => {
    trace.push(`pause ${milliseconds}`);
  };
  const audit = options.audit ?? createRiskyUsersAudit({ pause });
  const client = new GraphHttpClient({
    getAccessToken: async () => "synthetic",
    // The client's own retry back-off is not under test here.
    sleep: async () => {},
    maxResponseBytes: 8_388_608,
    fetchImplementation: async (input) => {
      const url = new URL(String(input));
      urls.push(url);
      trace.push(`GET ${url.pathname.replace("/v1.0/", "")}`);
      return route(url) ?? new Response("{}", { status: 404 });
    },
  });
  const context: AuditContext = {
    graph: scopeGraph(client, [...audit.packs, ...audit.optionalPacks]),
    now,
    granted: new Set(["IdentityRiskyUser.Read.All"]),
    signal: new AbortController().signal,
    scope: options.scope ?? null,
  };
  return { audit, context, trace, urls };
}

const listing =
  (users: unknown[]): Route =>
  (url) =>
    url.pathname === "/v1.0/identityProtection/riskyUsers"
      ? Response.json({ value: users })
      : undefined;

async function runWith(users: unknown[], options: Parameters<typeof setup>[1] = {}) {
  const { audit, context, trace, urls } = setup(listing(users), options);
  return { result: await audit.run(context), trace, urls };
}

const row = <T extends { key: string }>(rows: T[], key: string): T => {
  const found = rows.find((item) => item.key === key);
  if (!found) throw new Error(`No row ${key}`);
  return found;
};
const tile = (summary: { label: string; value: number }[], label: string) =>
  summary.find((item) => item.label === label);
const definition = auditDefinition(riskyUsersAudit);
const keysIn = (rows: Parameters<typeof queryAuditRows>[1], view: string) =>
  queryAuditRows(definition, rows, { view })
    .map((item) => item.key)
    .sort();

describe("the risky users audit definition", () => {
  it("is in the catalog, scoped, and needs only the risk pack", () => {
    expect(builtInAudits).toContain(riskyUsersAudit);
    expect(definition).toMatchObject({
      id: "identity.risky-users",
      category: "Identity",
      scoped: true,
      permissions: ["IdentityRiskyUser.Read.All"],
      optionalPermissions: [],
      defaultView: "at-risk",
    });
    expect(definition.views.map((view) => view.id)).toEqual([
      "at-risk",
      "high-risk",
      "closed",
      "all",
    ]);
    expect(definition.columns.some((column) => column.key === "domain")).toBe(true);
  });

  it("says Microsoft Entra ID P2 is needed when Microsoft refuses", () => {
    expect(riskyUsersAudit.forbidden).toContain(P2);
    expect(riskyUsersAudit.forbidden).toContain("IdentityRiskyUser.Read.All");
    expect(riskyUsersAudit.forbidden).toContain("Security Operator");
  });

  it("paces real runs at more than one second between pages", () => {
    expect(REQUEST_SPACING_MILLISECONDS).toBeGreaterThan(1_000);
  });
});

describe("rows", () => {
  it("has one valid row per risky user, linked to the user's page", async () => {
    const { result } = await runWith(tenant);
    expect(result.rows).toHaveLength(tenant.length);
    for (const item of result.rows) auditRowSchema.parse(item);
    expect(row(result.rows, "high-open")).toMatchObject({
      personId: "high-open",
      cells: {
        person: { name: "User high-open", principal: "high-open@example.invalid" },
        level: "High",
        state: "At risk",
        detail: "None",
        updated: daysAgo(2),
        days: 2,
        processing: false,
        domain: "example.invalid",
      },
    });
    expect(result.complete).toBe(true);
  });

  it("reads every risk level, state and detail as words", async () => {
    const { result } = await runWith(tenant);
    const cells = (key: string) => row(result.rows, key).cells;
    expect(cells("medium-open").level).toBe("Medium");
    expect(cells("low-open").level).toBe("Low");
    expect(cells("uncovered")).toMatchObject({ level: "Hidden", state: "None", detail: "Hidden" });
    expect(cells("compromised")).toMatchObject({
      state: "Confirmed compromised",
      detail: "Admin confirmed the user was compromised",
      processing: true,
    });
    expect(cells("fixed-recently")).toMatchObject({
      level: "None",
      state: "Remediated",
      detail: "User reset the password securely",
    });
    expect(cells("dismissed-old")).toMatchObject({
      state: "Dismissed",
      detail: "Admin dismissed all risk for the user",
    });
    expect(cells("safe")).toMatchObject({
      state: "Confirmed safe",
      detail: "Admin confirmed the sign-in was safe",
    });
  });

  it("falls back to plain words for values Microsoft may add later", async () => {
    const { result } = await runWith([
      risky("future", {
        riskLevel: "unknownFutureValue",
        riskState: "somethingNew",
        riskDetail: "adminDidSomethingNew",
      }),
    ]);
    expect(row(result.rows, "future").cells).toMatchObject({
      level: "Unknown",
      state: "Unknown",
      detail: "Admin did something new",
    });
  });

  it("counts days since the update and never a negative number", async () => {
    const { result } = await runWith([
      risky("a", { riskLastUpdatedDateTime: daysAgo(40) }),
      risky("b", { riskLastUpdatedDateTime: new Date(now.getTime() + 3_600_000).toISOString() }),
      risky("c", { riskLastUpdatedDateTime: null }),
    ]);
    expect(row(result.rows, "a").cells.days).toBe(40);
    expect(row(result.rows, "b").cells.days).toBe(0);
    expect(row(result.rows, "c").cells).toMatchObject({ days: null, updated: null });
  });
});

describe("views", () => {
  it("separates users still at risk, high risk, and closed", async () => {
    const { result } = await runWith(tenant);
    expect(keysIn(result.rows, "at-risk")).toEqual([
      "compromised",
      "high-open",
      "low-open",
      "medium-open",
    ]);
    expect(keysIn(result.rows, "high-risk")).toEqual(["compromised", "high-open"]);
    expect(keysIn(result.rows, "closed")).toEqual(["dismissed-old", "fixed-recently"]);
    expect(keysIn(result.rows, "all")).toHaveLength(tenant.length);
  });

  it("leaves a high-risk user who was dismissed out of the high risk view", async () => {
    const { result } = await runWith([
      risky("settled", { riskLevel: "high", riskState: "dismissed" }),
      risky("open"),
    ]);
    expect(keysIn(result.rows, "high-risk")).toEqual(["open"]);
    expect(keysIn(result.rows, "closed")).toEqual(["settled"]);
  });
});

describe("summary tiles", () => {
  it("counts each group and links it to its view", async () => {
    const { result } = await runWith(tenant);
    expect(result.summary.map((item) => auditSummaryItemSchema.parse(item))).toEqual([
      { label: "Confirmed compromised", value: 1, tone: "critical", view: "at-risk" },
      { label: "High risk, still at risk", value: 2, tone: "critical", view: "high-risk" },
      { label: "Medium risk, still at risk", value: 1, tone: "warning", view: "at-risk" },
      {
        label: "Remediated or dismissed in the last 30 days",
        value: 1,
        tone: "neutral",
        view: "closed",
      },
    ]);
  });

  it("counts a confirmed compromised user in the high risk tile only when the level is high", async () => {
    const { result } = await runWith([
      risky("c", { riskLevel: "medium", riskState: "confirmedCompromised" }),
    ]);
    expect(tile(result.summary, "Confirmed compromised")?.value).toBe(1);
    expect(tile(result.summary, "High risk, still at risk")?.value).toBe(0);
    expect(tile(result.summary, "Medium risk, still at risk")?.value).toBe(1);
  });

  it("leaves a high-risk user who has been dismissed out of the still-at-risk tile", async () => {
    const { result } = await runWith([
      risky("settled", { riskLevel: "high", riskState: "dismissed" }),
      risky("remediated", { riskLevel: "high", riskState: "remediated" }),
      risky("open"),
    ]);
    expect(tile(result.summary, "High risk, still at risk")?.value).toBe(1);
  });

  it("is calm when nobody is at risk", async () => {
    const { result } = await runWith([
      risky("safe", { riskLevel: "none", riskState: "confirmedSafe" }),
    ]);
    expect(tile(result.summary, "Confirmed compromised")).toMatchObject({ value: 0, tone: "good" });
    expect(tile(result.summary, "High risk, still at risk")).toMatchObject({
      value: 0,
      tone: "good",
    });
    expect(tile(result.summary, "Medium risk, still at risk")).toMatchObject({
      value: 0,
      tone: "neutral",
    });
  });

  it("counts a closed user only when it was closed in the last 30 days", async () => {
    const closed = (id: string, days: number, state = "remediated") =>
      risky(id, { riskLevel: "none", riskState: state, riskLastUpdatedDateTime: daysAgo(days) });
    const { result } = await runWith([
      closed("a", 0),
      closed("b", 29, "dismissed"),
      closed("c", 30),
      closed("d", 31),
      closed("e", 400, "dismissed"),
      risky("open", { riskLastUpdatedDateTime: daysAgo(1) }),
      risky("unknown-date", {
        riskLevel: "none",
        riskState: "remediated",
        riskLastUpdatedDateTime: null,
      }),
    ]);
    expect(tile(result.summary, "Remediated or dismissed in the last 30 days")?.value).toBe(3);
  });
});

describe("deleted users", () => {
  it("skips them and says how many", async () => {
    const { result } = await runWith([
      ...tenant,
      risky("gone-1", { isDeleted: true }),
      risky("gone-2", { isDeleted: true, riskState: "confirmedCompromised" }),
    ]);
    expect(result.rows.map((item) => item.key)).not.toContain("gone-1");
    expect(result.rows).toHaveLength(tenant.length);
    expect(result.notes.join(" ")).toContain(
      "2 deleted users that Microsoft still lists were left out",
    );
    // A deleted compromised user is not a compromised user to act on.
    expect(tile(result.summary, "Confirmed compromised")?.value).toBe(1);
    expect(result.complete).toBe(true);
  });

  it("says one deleted user in the singular", async () => {
    const { result } = await runWith([risky("gone", { isDeleted: true })]);
    expect(result.notes.join(" ")).toContain(
      "1 deleted user that Microsoft still lists was left out",
    );
  });

  it("counts only the deleted users the scope covers", async () => {
    const { result } = await runWith(
      [
        risky("kept", { userPrincipalName: "a@kept.example.invalid" }),
        risky("gone-inside", { isDeleted: true, userPrincipalName: "b@kept.example.invalid" }),
        risky("gone-outside", { isDeleted: true, userPrincipalName: "c@other.example.invalid" }),
      ],
      { scope: new Set(["kept.example.invalid"]) },
    );
    expect(result.notes.join(" ")).toContain(
      "1 deleted user that Microsoft still lists was left out",
    );
    // A deleted user in another domain is not an account the scope left out either.
    expect(result.notes.join(" ")).not.toContain("accounts in other domains");
  });

  it("does not offer a deleted user's domain for the scope", async () => {
    const { result } = await runWith([
      risky("kept"),
      risky("gone", { isDeleted: true, userPrincipalName: "gone@retired.example.invalid" }),
    ]);
    expect(result.domains).toEqual([{ domain: "example.invalid", accounts: 1 }]);
  });
});

describe("the audit scope", () => {
  const mixed = [
    risky("inside-open", { userPrincipalName: "a@kept.example.invalid" }),
    risky("inside-closed", {
      userPrincipalName: "b@kept.example.invalid",
      riskLevel: "none",
      riskState: "remediated",
    }),
    risky("outside-high", { userPrincipalName: "c@other.example.invalid" }),
    risky("outside-medium", { userPrincipalName: "d@other.example.invalid", riskLevel: "medium" }),
    risky("outside-compromised", {
      userPrincipalName: "e@third.example.invalid",
      riskLevel: "low",
      riskState: "confirmedCompromised",
    }),
    risky("no-principal", { userPrincipalName: null }),
  ];
  const scope = new Set(["kept.example.invalid"]);

  it("keeps only the accounts in the chosen domains", async () => {
    const { result } = await runWith(mixed, { scope });
    expect(result.rows.map((item) => item.key).sort()).toEqual(["inside-closed", "inside-open"]);
    expect(result.notes[0]).toBe("Limited to accounts in: kept.example.invalid.");
    expect(result.notes[1]).toBe("4 accounts in other domains were left out by the audit scope.");
  });

  it("computes the tiles from the accounts in scope", async () => {
    const { result } = await runWith(mixed, { scope });
    expect(tile(result.summary, "High risk, still at risk")?.value).toBe(1);
    expect(tile(result.summary, "Medium risk, still at risk")?.value).toBe(0);
    expect(tile(result.summary, "Confirmed compromised")).toMatchObject({
      value: 0,
      tone: "neutral",
    });
  });

  it("lists every domain it read, whatever the scope", async () => {
    const { result } = await runWith(mixed, { scope });
    expect(result.domains).toEqual([
      { domain: "kept.example.invalid", accounts: 2 },
      { domain: "other.example.invalid", accounts: 2 },
      { domain: "third.example.invalid", accounts: 1 },
    ]);
  });

  it("never lets the scope hide a high-risk or compromised user silently", async () => {
    const { result } = await runWith(mixed, { scope });
    expect(tile(result.summary, "High risk or compromised outside the scope")).toMatchObject({
      // The high-risk and the compromised user outside the scope, and the high-risk user with no
      // principal name, which a scope never covers.
      value: 3,
      tone: "critical",
      view: null,
    });
    expect(result.notes.join(" ")).toContain(
      "3 users at high risk or confirmed compromised are in domains the audit scope leaves out",
    );
  });

  it("shows no outside tile when nothing urgent is outside", async () => {
    const { result } = await runWith(
      [risky("in", { userPrincipalName: "a@kept.example.invalid" }), mixed[3]],
      { scope },
    );
    expect(tile(result.summary, "High risk or compromised outside the scope")).toBeUndefined();
  });

  it("covers every account when there is no scope", async () => {
    const { result } = await runWith(mixed);
    expect(result.rows).toHaveLength(mixed.length);
    expect(result.notes.join(" ")).not.toContain("Limited to");
  });
});

describe("when Microsoft refuses or has nothing", () => {
  it("passes a 403 up so the P2 explanation is shown", async () => {
    const { audit, context } = setup(() => new Response("{}", { status: 403 }));
    await expect(audit.run(context)).rejects.toMatchObject({
      constructor: GraphRequestError,
      status: 403,
    });
    expect(audit.forbidden).toContain(P2);
  });

  it("says Entra ID P2 is needed when the list is empty", async () => {
    const { result } = await runWith([]);
    expect(result.rows).toEqual([]);
    expect(result.notes.join(" ")).toContain(P2);
    expect(result.notes.join(" ")).toContain("Microsoft returned no risky users");
    expect(result.complete).toBe(true);
  });

  it("does not say it when users came back, even if none are shown", async () => {
    const { result } = await runWith([risky("gone", { isDeleted: true })]);
    expect(result.notes.join(" ")).not.toContain(P2);
  });

  it("does not say it when the scope leaves everyone out", async () => {
    const { result } = await runWith([risky("elsewhere")], {
      scope: new Set(["kept.example.invalid"]),
    });
    expect(result.rows).toEqual([]);
    expect(result.notes.join(" ")).not.toContain(P2);
  });
});

describe("records Microsoft sent in a shape Prism365 cannot read", () => {
  it("leaves them out, counts them and marks the result partial", async () => {
    const { result } = await runWith([
      risky("fine"),
      { riskLevel: "high", userPrincipalName: "no-id@example.invalid" },
      risky("odd id", { id: "has spaces and ! marks" }),
      "not an object",
    ]);
    expect(result.rows.map((item) => item.key)).toEqual(["fine"]);
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("3 users could not be read and are left out");
  });

  it("tolerates nulls, wrong types and unknown fields", async () => {
    const { result } = await runWith([
      {
        id: "sparse",
        isDeleted: null,
        isProcessing: "yes",
        riskLastUpdatedDateTime: "yesterday",
        riskLevel: null,
        riskState: 7,
        riskDetail: null,
        userDisplayName: null,
        userPrincipalName: "sparse\u0000@example.invalid",
        somethingNew: { nested: true },
      },
    ]);
    expect(row(result.rows, "sparse").cells).toMatchObject({
      person: { name: null, principal: "sparse@example.invalid" },
      level: "Unknown",
      state: "Unknown",
      detail: null,
      updated: null,
      days: null,
      processing: null,
    });
    expect(result.complete).toBe(true);
  });

  it("keeps every note inside the length a run stores", async () => {
    const { result } = await runWith([...tenant, risky("gone", { isDeleted: true }), "bad"], {
      scope: new Set(["example.invalid"]),
    });
    for (const note of result.notes) {
      expect(note.length).toBeLessThanOrEqual(500);
      expect(note).not.toMatch(/[—·!]/);
    }
  });
});

describe("paging and pacing", () => {
  const base = "https://graph.microsoft.com/v1.0/identityProtection/riskyUsers";

  it("asks for 500 users a page", async () => {
    const { urls } = await runWith(tenant);
    expect(urls).toHaveLength(1);
    expect(urls[0]?.searchParams.get("$top")).toBe("500");
  });

  it("waits between pages and not before the first, with no real waiting", async () => {
    const { audit, context, trace } = setup((url) => {
      const page = Number(url.searchParams.get("$skiptoken") ?? "0");
      return Response.json({
        value: [risky(`page-${page}`)],
        ...(page < 2 ? { "@odata.nextLink": `${base}?$top=500&$skiptoken=${page + 1}` } : {}),
      });
    });
    const result = await audit.run(context);
    expect(result.rows.map((item) => item.key).sort()).toEqual(["page-0", "page-1", "page-2"]);
    expect(trace).toEqual([
      "GET identityProtection/riskyUsers",
      `pause ${REQUEST_SPACING_MILLISECONDS}`,
      "GET identityProtection/riskyUsers",
      `pause ${REQUEST_SPACING_MILLISECONDS}`,
      "GET identityProtection/riskyUsers",
    ]);
    expect(result.complete).toBe(true);
  });

  it("stops at its page limit and says the result is partial", async () => {
    const { audit, context, urls } = setup((url) => {
      const page = Number(url.searchParams.get("$skiptoken") ?? "0");
      return Response.json({
        value: [risky(`page-${page}`)],
        "@odata.nextLink": `${base}?$top=500&$skiptoken=${page + 1}`,
      });
    });
    const result = await audit.run(context);
    expect(urls).toHaveLength(25);
    expect(result.rows).toHaveLength(25);
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("Stopped at 10,000 risky users");
  });

  it("stops waiting when the audit is stopped", async () => {
    const controller = new AbortController();
    const audit = createRiskyUsersAudit({
      pause: async (_milliseconds, signal) => {
        controller.abort(new Error("time limit"));
        signal.throwIfAborted();
      },
    });
    const { context, urls } = setup(
      () => Response.json({ value: [risky("a")], "@odata.nextLink": `${base}?$skiptoken=1` }),
      { audit },
    );
    await expect(audit.run({ ...context, signal: controller.signal })).rejects.toThrow(
      "time limit",
    );
    expect(urls).toHaveLength(1);
  });

  it("refuses a continuation link outside the risky users endpoint", async () => {
    const { audit, context } = setup(() =>
      Response.json({
        value: [risky("a")],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/users?$skiptoken=1",
      }),
    );
    await expect(audit.run(context)).rejects.toMatchObject({ reason: "endpoint" });
  });
});
