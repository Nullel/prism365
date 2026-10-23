import { queryAuditRows } from "@prism365/contracts";
import { GraphRetryDeferredError } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import {
  auditDefinition,
  createM365AppUsageAudit,
  lastActivityOf,
  MAX_USAGE_ACCOUNTS,
  m365AppUsageAudit,
  usageOf,
} from "../src/index.ts";
import {
  expectValidTable,
  fake,
  find,
  now,
  paged,
  type Route,
} from "./apps-devices-usage-support.ts";

// Synthetic accounts on a.example and b.example. The report names people by sign-in name only.
const audit = createM365AppUsageAudit({ pageDelayMilliseconds: 0 });
const REPORT = "/v1.0/reports/getM365AppUserDetail(period='D90')";
const SETTINGS = "/v1.0/admin/reportSettings";
const granted = ["Reports.Read.All", "ReportSettings.Read.All"];
const grantedWithoutSettings = ["Reports.Read.All"];

const flags = {
  windows: false,
  mac: false,
  mobile: false,
  web: false,
  outlook: false,
  word: false,
  excel: false,
  powerPoint: false,
  oneNote: false,
  teams: false,
};
const record = (
  userPrincipalName: string,
  lastActivityDate: string | null,
  used: Partial<typeof flags> = {},
  extra: Record<string, unknown> = {},
) => ({
  reportRefreshDate: "2026-10-20",
  userPrincipalName,
  lastActivationDate: "2025-03-01",
  lastActivityDate,
  details: [{ reportPeriod: 90, ...flags, ...used }],
  ...extra,
});

const records = [
  record("ada@a.example", "2026-10-20", {
    windows: true,
    outlook: true,
    word: true,
    teams: true,
  }),
  record("ben@a.example", "2026-10-21", { teams: true, web: true, mobile: true }),
  record("cy@b.example", "2026-10-15", { outlook: true, web: true }),
  record("dee@a.example", ""),
  record("eli@b.example", null),
  record("fay@a.example", "2026-10-01", { mac: true, powerPoint: true, teams: true }),
  record("gus@a.example", "2026-06-01", { windows: true, excel: true }),
  // The flat shape, with the flags on the record itself.
  {
    userPrincipalName: "hal@a.example",
    lastActivityDate: "2026-10-21",
    windows: true,
    outlook: true,
  },
  record("ivy@a.example", "yesterday", { outlook: true }),
  record("jo@b.example", "2026-10-10", { mobile: true, outlook: true }),
  // Unreadable: no sign-in name.
  { displayName: "no principal", lastActivityDate: "2026-10-01" },
];

const tenant: Route = (url) => {
  if (url.pathname === REPORT) return paged(url, records, 200);
  if (url.pathname === SETTINGS) return Response.json({ displayConcealedNames: false });
  return undefined;
};

const run = (route: Route = tenant, grantedScopes: string[] = granted, scope?: string[]) =>
  audit.run(
    fake(audit, route, { granted: grantedScopes, scope: scope ? new Set(scope) : null }).context,
  );
const note = (result: { notes: string[] }) => result.notes.join("\n");

// Concealed identifiers are 32 hexadecimal characters.
const hashed = records
  .filter((item) => "userPrincipalName" in item)
  .map((item, index) => ({
    ...item,
    userPrincipalName: index.toString(16).padStart(32, "0"),
  }));
const concealedTenant = (setting: Response | null): Route => {
  return (url) => {
    if (url.pathname === REPORT) return paged(url, hashed, 200);
    if (url.pathname === SETTINGS) return setting ?? undefined;
    return undefined;
  };
};

describe("Microsoft 365 app usage", () => {
  it("publishes a valid table", async () => {
    await expectValidTable(audit, tenant, { granted });
  });

  it("declares the usage pack, with the privacy setting as an optional pack, and a scope", () => {
    expect(auditDefinition(audit)).toMatchObject({
      id: "usage.m365-apps",
      category: "Licenses",
      permissions: ["Reports.Read.All"],
      optionalPermissions: ["ReportSettings.Read.All"],
      scoped: true,
    });
  });

  describe("with names visible", () => {
    it("lists each account with its last activity, apps and platforms", async () => {
      const result = await run();
      expect(result.rows).toHaveLength(10);
      expect(find(result.rows, "ada@a.example")).toMatchObject({
        personId: null,
        cells: {
          person: { name: null, principal: "ada@a.example" },
          lastActivity: "2026-10-20T00:00:00.000Z",
          days: 2,
          apps: ["Outlook", "Word", "Teams"],
          platforms: ["Windows"],
          outlook: true,
          word: true,
          excel: false,
          powerPoint: false,
          oneNote: false,
          teams: true,
          domain: "a.example",
          activated: "2025-03-01T00:00:00.000Z",
        },
      });
      expect(find(result.rows, "ben@a.example").cells).toMatchObject({
        apps: ["Teams"],
        platforms: ["Mobile", "Web"],
      });
    });

    it("reads both the nested and the flat shape", async () => {
      const result = await run();
      expect(find(result.rows, "hal@a.example").cells).toMatchObject({
        apps: ["Outlook"],
        platforms: ["Windows"],
        outlook: true,
      });
      expect(find(result.rows, "hal@a.example").tags).toEqual(["desktop"]);
    });

    it("tags accounts with no activity in the period, including a date older than the period", async () => {
      const result = await run();
      const idle = result.rows.filter((row) => row.tags.includes("idle"));
      expect(idle.map((row) => row.key).sort()).toEqual([
        "dee@a.example",
        "eli@b.example",
        "gus@a.example",
      ]);
      expect(find(result.rows, "dee@a.example").cells).toMatchObject({
        lastActivity: null,
        days: null,
        apps: [],
      });
      expect(find(result.rows, "gus@a.example").cells.days).toBe(143);
    });

    it("never counts a date it cannot read as idle, and says the result is partial", async () => {
      const result = await run();
      expect(find(result.rows, "ivy@a.example")).toMatchObject({
        tags: [],
        cells: { lastActivity: null, days: null },
      });
      expect(note(result)).toContain("1 account has a last activity date Prism365 could not read");
      expect(result.complete).toBe(false);
    });

    it("opens each view on the accounts that belong in it", async () => {
      const result = await run();
      const definition = auditDefinition(audit);
      const keys = (view: string) =>
        queryAuditRows(definition, result.rows, { view })
          .map((row) => row.key)
          .sort();
      expect(keys("idle")).toEqual(["dee@a.example", "eli@b.example", "gus@a.example"]);
      expect(keys("teams-only")).toEqual(["ben@a.example"]);
      expect(keys("web-only")).toEqual(["cy@b.example"]);
      expect(keys("desktop")).toEqual([
        "ada@a.example",
        "fay@a.example",
        "gus@a.example",
        "hal@a.example",
      ]);
      expect(keys("web-mobile")).toEqual(["ben@a.example", "cy@b.example", "jo@b.example"]);
      expect(keys("all")).toHaveLength(10);
      expect(definition.defaultView).toBe("idle");
    });

    it("summarises with the tiles in the contract's order, each opening its view", async () => {
      const result = await run();
      expect(result.summary).toEqual([
        {
          label: "No Microsoft 365 app activity in 90 days",
          value: 3,
          tone: "warning",
          view: "idle",
        },
        { label: "Desktop apps used", value: 4, tone: "neutral", view: "desktop" },
        { label: "Web and mobile only", value: 3, tone: "neutral", view: "web-mobile" },
      ]);
    });

    it("shows a tenant where everyone is active as good", async () => {
      const result = await run((url) =>
        url.pathname === REPORT
          ? paged(url, [record("ada@a.example", "2026-10-20", { windows: true, outlook: true })])
          : tenant(url),
      );
      expect(result.summary[0]).toMatchObject({ value: 0, tone: "good" });
      expect(result.complete).toBe(true);
    });

    it("says the rows cannot open a person page, and why", async () => {
      const text = note(await run());
      expect(text).toContain("Rows do not link to a person page");
      expect(text).toContain("sign-in name, not by object ID");
    });

    it("carries the report's limits in its notes", async () => {
      const text = note(await run());
      expect(text).toContain("lag by 24 to 72 hours");
      expect(text).toContain("last 90 days");
      expect(text).toContain("Reports Reader");
      expect(text).toContain("Global Reader sees tenant totals only");
      expect(text).toContain("does not measure anyone's work");
    });

    it("leaves out a record with no sign-in name and says so", async () => {
      const result = await run();
      expect(note(result)).toContain("1 account in the report could not be read and is left out");
      expect(
        result.rows.some(
          (row) =>
            row.cells.lastActivity === "2026-10-01T00:00:00.000Z" && row.key === "no principal",
        ),
      ).toBe(false);
    });
  });

  describe("the request", () => {
    it("reads the 90-day Microsoft 365 apps report as JSON, then the privacy setting", async () => {
      const ctx = fake(audit, tenant, { granted });
      await audit.run(ctx.context);
      expect(ctx.requests.map((url) => url.pathname)).toEqual([SETTINGS, REPORT]);
      const read = ctx.requests[1];
      expect(read?.searchParams.get("$format")).toBe("application/json");
      expect(read?.searchParams.get("$top")).toBe("1000");
    });

    it("does not read the privacy setting without its permission", async () => {
      const ctx = fake(audit, tenant, { granted: grantedWithoutSettings });
      await audit.run(ctx.context);
      expect(ctx.requests.map((url) => url.pathname)).toEqual([REPORT]);
    });

    it("falls back to Microsoft's default page when a larger one is rejected", async () => {
      const route: Route = (url) =>
        url.pathname === REPORT && url.searchParams.has("$top")
          ? new Response("{}", { status: 400 })
          : tenant(url);
      const ctx = fake(audit, route, { granted: grantedWithoutSettings });
      const result = await audit.run(ctx.context);
      expect(result.rows).toHaveLength(10);
      expect(ctx.requests).toHaveLength(2);
      expect(ctx.requests[1]?.searchParams.has("$top")).toBe(false);
      expect(ctx.requests[1]?.searchParams.get("$format")).toBe("application/json");
    });

    it("follows every page of the report", async () => {
      const ctx = fake(
        audit,
        (url) => (url.pathname === REPORT ? paged(url, records, 4) : tenant(url)),
        { granted: grantedWithoutSettings },
      );
      const result = await audit.run(ctx.context);
      expect(result.rows).toHaveLength(10);
      expect(ctx.requests).toHaveLength(3);
      // The continuation keeps the JSON format.
      expect(ctx.requests[2]?.searchParams.get("$format")).toBe("application/json");
    });

    it("stops at 10,000 accounts and says so", async () => {
      const many = Array.from({ length: MAX_USAGE_ACCOUNTS + 300 }, (_, index) =>
        record(`user${index}@a.example`, "2026-10-20", { web: true, outlook: true }),
      );
      const result = await run((url) =>
        url.pathname === REPORT ? paged(url, many, 1_000) : tenant(url),
      );
      expect(result.rows).toHaveLength(MAX_USAGE_ACCOUNTS);
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("Stopped at 10,000 accounts");
    });
  });

  describe("pacing", () => {
    const pages = (url: URL) => (url.pathname === REPORT ? paged(url, records, 4) : tenant(url));

    it("pauses between pages, not before the first", async () => {
      const slow = createM365AppUsageAudit({ pageDelayMilliseconds: 40 });
      const started = performance.now();
      await slow.run(fake(slow, pages, { granted: grantedWithoutSettings }).context);
      const elapsed = performance.now() - started;
      // Three pages, so two pauses.
      expect(elapsed).toBeGreaterThanOrEqual(70);
      expect(elapsed).toBeLessThan(1_000);
    });

    it("paces the built-in audit by default", async () => {
      const twoPages = (url: URL) =>
        url.pathname === REPORT ? paged(url, records, 6) : tenant(url);
      const started = performance.now();
      await m365AppUsageAudit.run(
        fake(m365AppUsageAudit, twoPages, { granted: grantedWithoutSettings }).context,
      );
      expect(performance.now() - started).toBeGreaterThanOrEqual(400);
    });

    it("stops waiting when the run is stopped", async () => {
      const slow = createM365AppUsageAudit({ pageDelayMilliseconds: 60_000 });
      const controller = new AbortController();
      const ctx = fake(slow, pages, { granted: grantedWithoutSettings, signal: controller.signal });
      const started = performance.now();
      const running = slow.run(ctx.context);
      setTimeout(() => controller.abort(new Error("stopped")), 30);
      await expect(running).rejects.toThrow("stopped");
      expect(performance.now() - started).toBeLessThan(2_000);
    });
  });

  describe("when Microsoft pushes back", () => {
    it("keeps the pages it read when throttled later, and says how to finish", async () => {
      const route: Route = (url) => {
        if (url.pathname !== REPORT) return tenant(url);
        return url.searchParams.has("$skiptoken")
          ? new Response("{}", { status: 429 })
          : paged(url, records, 4);
      };
      const result = await run(route, grantedWithoutSettings);
      expect(result.rows).toHaveLength(4);
      expect(result.complete).toBe(false);
      expect(note(result)).toContain("asked Prism365 to slow down after 4 accounts");
      expect(note(result)).toContain("Run the audit again in about ten minutes");
    });

    it("fails when it is throttled before reading anything", async () => {
      await expect(
        run((url) => (url.pathname === REPORT ? new Response("{}", { status: 429 }) : undefined)),
      ).rejects.toBeInstanceOf(GraphRetryDeferredError);
    });

    it("fails with the refusal, which names the roles that can read per-user usage", async () => {
      await expect(
        run((url) => (url.pathname === REPORT ? new Response("{}", { status: 403 }) : tenant(url))),
      ).rejects.toMatchObject({ status: 403 });
      const text = audit.forbidden;
      expect(text).toContain("Reports.Read.All");
      expect(text).toContain("Reports Reader");
      expect(text).toContain("Global Reader sees tenant totals only");
    });
  });

  describe("the privacy setting", () => {
    it("shows names when Microsoft says it is off", async () => {
      const result = await run();
      expect(result.rows).toHaveLength(10);
      expect(note(result)).not.toContain("privacy setting");
    });

    it("reads the setting from either place Microsoft documents it", async () => {
      const nested = await run(
        concealedRoute(Response.json({ value: { displayConcealedNames: true } })),
      );
      expect(nested.rows).toEqual([]);
      const flat = await run(concealedRoute(Response.json({ displayConcealedNames: true })));
      expect(flat.rows).toEqual([]);
    });

    function concealedRoute(setting: Response): Route {
      return (url) => {
        if (url.pathname === SETTINGS) return setting;
        return tenant(url);
      };
    }

    it("shows names, and says why, when the setting cannot be read but names are readable", async () => {
      const refused = await run((url) =>
        url.pathname === SETTINGS ? new Response("{}", { status: 403 }) : tenant(url),
      );
      expect(refused.rows).toHaveLength(10);
      expect(note(refused)).toContain("Microsoft refused the report privacy setting");
      expect(note(refused)).toContain(
        "Names are shown because Microsoft returned readable sign-in names",
      );
      const absent = await run(tenant, grantedWithoutSettings);
      expect(absent.rows).toHaveLength(10);
      expect(note(absent)).toContain("it needs ReportSettings.Read.All");
    });

    it("fails the run on other errors reading the setting", async () => {
      await expect(
        run((url) =>
          url.pathname === SETTINGS ? new Response("{}", { status: 500 }) : tenant(url),
        ),
      ).rejects.toMatchObject({ status: 500 });
    });
  });

  describe("with names concealed", () => {
    it("gives totals and no rows, and never unmasks anything", async () => {
      const result = await run(concealedTenant(Response.json({ displayConcealedNames: true })));
      expect(result.rows).toEqual([]);
      expect(result.summary).toEqual([
        {
          label: "No Microsoft 365 app activity in 90 days",
          value: 3,
          tone: "warning",
          view: null,
        },
        { label: "Desktop apps used", value: 4, tone: "neutral", view: null },
        { label: "Web and mobile only", value: 3, tone: "neutral", view: null },
      ]);
      // Nothing from the report's identifiers reaches the result.
      expect(JSON.stringify(result)).not.toMatch(/[0-9a-f]{32}/);
      expect(result.domains).toBeUndefined();
      expect(note(result)).toContain(
        "Names are concealed in this tenant's Microsoft usage reports",
      );
      expect(note(result)).toContain("Global Administrator can turn off concealment");
      expect(note(result)).toContain("Settings, Org settings, Reports");
      expect(note(result)).toContain("will not unmask anything itself");
      expect(note(result)).not.toContain("Rows do not link");
    });

    it("treats hashed identifiers as concealed when the setting cannot be read", async () => {
      const result = await run(concealedTenant(null), grantedWithoutSettings);
      expect(result.rows).toEqual([]);
      expect(result.summary[0]).toMatchObject({ value: 3, view: null });
      expect(note(result)).toContain("could not read the report privacy setting");
      expect(note(result)).toContain("look like hashes");
      expect(note(result)).toContain("Global Administrator can turn off concealment");
    });

    it("keeps them concealed when the setting says names are shown but the report is still hashed", async () => {
      const result = await run(concealedTenant(Response.json({ displayConcealedNames: false })));
      expect(result.rows).toEqual([]);
      expect(note(result)).toContain(
        "setting says names are shown, but the report still returns hashed",
      );
    });

    it("does not mistake a few hash-like names for concealment", async () => {
      const mixed = [record("ada@a.example", "2026-10-20", { web: true }), ...hashed.slice(0, 2)];
      const result = await run((url) =>
        url.pathname === REPORT ? paged(url, mixed, 200) : tenant(url),
      );
      expect(result.rows).toHaveLength(3);
    });

    it("says an audit scope could not be applied", async () => {
      const result = await run(
        concealedTenant(Response.json({ displayConcealedNames: true })),
        granted,
        ["a.example"],
      );
      expect(note(result)).toContain("The audit scope could not be applied");
      expect(note(result)).not.toContain("Limited to accounts in");
      expect(result.summary[0]?.value).toBe(3);
    });
  });

  describe("the audit scope", () => {
    it("keeps only the chosen domains and counts only them", async () => {
      const result = await run(tenant, granted, ["b.example"]);
      expect(result.rows.map((row) => row.key).sort()).toEqual([
        "cy@b.example",
        "eli@b.example",
        "jo@b.example",
      ]);
      expect(result.summary).toEqual([
        {
          label: "No Microsoft 365 app activity in 90 days",
          value: 1,
          tone: "warning",
          view: "idle",
        },
        { label: "Desktop apps used", value: 0, tone: "neutral", view: "desktop" },
        { label: "Web and mobile only", value: 2, tone: "neutral", view: "web-mobile" },
      ]);
      expect(result.notes[0]).toBe("Limited to accounts in: b.example.");
      expect(note(result)).toContain(
        "7 accounts in other domains were left out by the audit scope",
      );
    });

    it("still lists every domain it read, so the chooser keeps offering them", async () => {
      const result = await run(tenant, granted, ["b.example"]);
      expect(result.domains).toEqual([
        { domain: "a.example", accounts: 7 },
        { domain: "b.example", accounts: 3 },
      ]);
    });

    it("covers everything with no scope", async () => {
      const result = await run();
      expect(result.rows).toHaveLength(10);
      expect(note(result)).not.toContain("Limited to accounts in");
    });
  });

  describe("how an account's use is read", () => {
    const usage = (used: Partial<typeof flags>) =>
      usageOf({ ...flags, ...used, userPrincipalName: "x@a.example", details: null } as never);

    it("calls an account Teams only when Teams is the one app used", () => {
      expect(usage({ teams: true })).toMatchObject({ teamsOnly: true });
      expect(usage({ teams: true, outlook: true })).toMatchObject({ teamsOnly: false });
      expect(usage({ outlook: true })).toMatchObject({ teamsOnly: false });
      expect(usage({})).toMatchObject({ teamsOnly: false, apps: [], platforms: [] });
    });

    it("calls an account web only when the browser is the only platform", () => {
      expect(usage({ web: true, outlook: true })).toMatchObject({ webOnly: true, webMobile: true });
      expect(usage({ web: true, mobile: true })).toMatchObject({ webOnly: false, webMobile: true });
      expect(usage({ web: true, windows: true })).toMatchObject({
        webOnly: false,
        webMobile: false,
        desktop: true,
      });
      expect(usage({ mobile: true })).toMatchObject({ webOnly: false, webMobile: true });
      expect(usage({ mac: true })).toMatchObject({ desktop: true });
    });

    it("counts a flag on any entry of details", () => {
      const record = {
        ...flags,
        userPrincipalName: "x@a.example",
        details: [{ ...flags }, { ...flags, mac: true }],
      };
      expect(usageOf(record as never).desktop).toBe(true);
    });

    it("reads dates without a time, and refuses ones that cannot be real", () => {
      expect(lastActivityOf("2026-10-20", now)).toEqual({
        at: "2026-10-20T00:00:00.000Z",
        days: 2,
        state: "known",
      });
      expect(lastActivityOf("2026-10-20T08:00:00Z", now).state).toBe("known");
      expect(lastActivityOf("", now).state).toBe("none");
      expect(lastActivityOf(null, now).state).toBe("none");
      for (const bad of ["yesterday", "1970-01-01", "2030-01-01", "20-10-2026"])
        expect(lastActivityOf(bad, now), bad).toEqual({
          at: null,
          days: null,
          state: "unreadable",
        });
    });

    it("counts a day exactly 90 days old as idle, and 89 as active", async () => {
      const dayOf = (days: number) =>
        new Date(now.getTime() - days * 86_400_000).toISOString().slice(0, 10);
      const result = await run((url) =>
        url.pathname === REPORT
          ? paged(url, [
              record("old@a.example", dayOf(90), { web: true }),
              record("new@a.example", dayOf(89), { web: true }),
            ])
          : tenant(url),
      );
      expect(find(result.rows, "old@a.example").tags).toContain("idle");
      expect(find(result.rows, "new@a.example").tags).not.toContain("idle");
    });
  });
});
