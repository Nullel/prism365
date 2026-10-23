import { randomUUID } from "node:crypto";
import { GraphHttpClient } from "@prism365/graph";
import {
  type CollectionCredential,
  latestRuns,
  listAttention,
  ModuleRegistry,
  ModuleScheduler,
} from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  checkState,
  createMessageCenterModule,
  createServiceHealthModule,
  plainText,
  postSections,
  serviceHealthStatus,
} from "../src/index.ts";

vi.setConfig({ testTimeout: 30_000 });

// Synthetic tenant; incident and post identifiers are invented.
const tenantId = randomUUID();
const start = new Date("2026-10-22T12:00:00Z");
let clock: Date;
let database: PGliteModuleDatabase;
let scheduler: ModuleScheduler | undefined;
let opened = false;
let issues: Record<string, unknown>[];
let messages: Record<string, unknown>[];
let rejectFilters: boolean;
let requests: string[];

function issue(overrides: Record<string, unknown> = {}) {
  return {
    id: "EX100001",
    title: "Some users cannot send mail to external recipients",
    service: "Exchange Online",
    feature: "E-Mail and calendar access",
    featureGroup: "E-Mail timely delivery",
    impactDescription:
      "Some users in the tenant may be unable to send email to external recipients.",
    classification: "incident",
    status: "investigating",
    origin: "microsoft",
    isResolved: false,
    startDateTime: "2026-10-22T09:00:00Z",
    endDateTime: null,
    lastModifiedDateTime: "2026-10-22T11:30:00Z",
    posts: [
      {
        createdDateTime: "2026-10-22T11:30:00Z",
        postType: "regular",
        description: {
          contentType: "html",
          content:
            "<p>Title: Some users cannot send mail</p><p>User impact: Users may be unable to send email to external recipients.</p><p>Scope of impact: Some users served by the affected infrastructure.</p><p>Workaround: Users can resend messages after a short delay.</p><p>Current status: We are investigating.</p>",
        },
      },
    ],
    ...overrides,
  };
}

function message(overrides: Record<string, unknown> = {}) {
  return {
    id: "MC200001",
    title: "Action required: update the client before the cutoff",
    category: "planForChange",
    severity: "normal",
    tags: ["Admin impact"],
    services: ["Microsoft Teams"],
    isMajorChange: true,
    actionRequiredByDateTime: "2026-11-01T00:00:00Z",
    startDateTime: "2026-10-01T00:00:00Z",
    endDateTime: "2026-12-01T00:00:00Z",
    lastModifiedDateTime: "2026-10-20T00:00:00Z",
    body: { contentType: "html", content: "<p>Update the client before the cutoff.</p>" },
    ...overrides,
  };
}

async function fakeFetch(input: RequestInfo | URL): Promise<Response> {
  const url = new URL(String(input));
  requests.push(`${url.pathname}${url.search}`);
  const filter = url.searchParams.get("$filter");
  if (url.pathname.endsWith("/healthOverviews"))
    return Response.json({
      value: [
        { id: "Exchange", service: "Exchange Online", status: "serviceDegradation" },
        { id: "microsoftteams", service: "Microsoft Teams", status: "serviceOperational" },
      ],
    });
  if (url.pathname.endsWith("/issues")) {
    if (filter && rejectFilters) return new Response("{}", { status: 400 });
    const value = !filter
      ? issues
      : filter === "isResolved eq false"
        ? issues.filter((item) => item.isResolved === false)
        : issues.filter(
            (item) =>
              Date.parse(String(item.lastModifiedDateTime)) >=
              Date.parse(filter.split(" ge ")[1] ?? ""),
          );
    return Response.json({ value });
  }
  if (url.pathname.endsWith("/messages")) return Response.json({ value: messages });
  return new Response("{}", { status: 404 });
}

async function setup(granted = ["ServiceHealth.Read.All", "ServiceMessage.Read.All"]) {
  const registry = new ModuleRegistry()
    .register(createServiceHealthModule())
    .register(createMessageCenterModule());
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  opened = true;
  scheduler = new ModuleScheduler({
    enabledPacks: registry.list().flatMap((module) => [...module.packs]),
    registry,
    database,
    now: () => clock,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: fakeFetch,
        sleep: async () => {},
      }),
  });
  const credential: CollectionCredential = {
    kind: "delegated",
    tenantId,
    actorId: randomUUID(),
    grantedPermissions: async () => ["User.Read", ...granted],
    accessToken: async () => "synthetic-token",
  };
  scheduler.bind(credential);
}

async function poll(minutes = 31) {
  await scheduler?.tick();
  clock = new Date(clock.getTime() + minutes * 60_000);
}

const attention = (kind: string) =>
  database.tenant(tenantId, (sql) =>
    listAttention(sql, tenantId, { filter: "all", kinds: [kind] }),
  );

beforeEach(() => {
  clock = start;
  issues = [issue()];
  messages = [message()];
  rejectFilters = false;
  requests = [];
});
afterEach(async () => {
  await scheduler?.stop();
  scheduler = undefined;
  if (opened) await database.close();
  opened = false;
});

describe("service health collector", () => {
  it("records Microsoft's incident with its impact, workaround and link, and the tenant's services", async () => {
    await setup();
    await poll();
    const [item] = await attention("service.incident");
    expect(item).toMatchObject({
      ref: "EX100001",
      class: "incident",
      severity: "high",
      services: ["Exchange Online"],
      unread: true,
      link: "https://admin.microsoft.com/#/servicehealth/:/alerts/EX100001",
      provider: {
        status: "Investigating",
        labels: ["Incident", "E-Mail and calendar access"],
        active: true,
        workaround: "Users can resend messages after a short delay.",
      },
    });
    // Microsoft's own words about who is affected, never "everyone".
    expect(item?.provider.impact).toContain("Some users in the tenant");
    expect(item?.provider.impact).toContain("Scope of impact: Some users served");
    const status = await database.tenant(tenantId, async (sql) =>
      serviceHealthStatus(sql, tenantId, {
        // A minute after the check, well inside the 15-minute freshness window.
        now: new Date(start.getTime() + 60_000),
        collecting: true,
        runs: await latestRuns(sql, tenantId),
        health: { enabled: true, granted: true },
        messages: { enabled: true, granted: true },
      }),
    );
    expect(status).toMatchObject({ state: "ok", lastSuccess: start.toISOString() });
    expect(status.services.map((service) => service.name)).toEqual([
      "Exchange Online",
      "Microsoft Teams",
    ]);
    expect(status.incidents.map((incident) => incident.ref)).toEqual(["EX100001"]);
  });

  it("updates one item as Microsoft changes the incident, without a duplicate per refresh", async () => {
    await setup();
    await poll();
    await poll();
    let items = await attention("service.incident");
    expect(items).toHaveLength(1);
    expect(items[0]?.revision).toBe(1);
    issues = [
      issue({
        status: "serviceRestored",
        isResolved: true,
        lastModifiedDateTime: "2026-10-22T12:40:00Z",
      }),
    ];
    await poll();
    items = await attention("service.incident");
    expect(items).toHaveLength(1);
    expect(items[0]).toMatchObject({ revision: 2, severity: "info", provider: { active: false } });
    expect(items[0]?.changeNote).toContain("Resolved by Microsoft.");
    expect(items[0]?.changeNote).toContain("Status: Investigating to Service restored");
  });

  it("falls back to reading the whole list when Microsoft rejects the filters", async () => {
    rejectFilters = true;
    issues = [
      issue(),
      issue({ id: "EX100002", isResolved: true, lastModifiedDateTime: "2026-09-01T00:00:00Z" }),
    ];
    await setup();
    await poll();
    expect((await attention("service.incident")).map((item) => item.ref)).toEqual(["EX100001"]);
    expect(requests.some((path) => path === "/v1.0/admin/serviceAnnouncement/issues")).toBe(true);
  });

  it("keeps Message center deadlines and labels, and leaves informational posts out", async () => {
    messages = [
      message(),
      message({
        id: "MC200002",
        title: "Feature X is being retired",
        tags: ["Retirement"],
        actionRequiredByDateTime: null,
        isMajorChange: false,
        category: "stayInformed",
      }),
      message({
        id: "MC200003",
        title: "We are retiring the old admin page",
        tags: [],
        actionRequiredByDateTime: null,
        isMajorChange: true,
        category: "planForChange",
      }),
      message({
        id: "MC200004",
        title: "New icons in the app bar",
        tags: ["User impact"],
        actionRequiredByDateTime: null,
        isMajorChange: false,
        category: "stayInformed",
      }),
    ];
    await setup();
    await poll();
    const items = await attention("service.message");
    expect(items.map((item) => item.ref).sort()).toEqual(["MC200001", "MC200002", "MC200003"]);
    const action = items.find((item) => item.ref === "MC200001");
    expect(action).toMatchObject({
      class: "message.action",
      provider: { deadline: "2026-11-01T00:00:00.000Z" },
      link: "https://admin.microsoft.com/#/MessageCenter/:/messages/MC200001",
    });
    expect(action?.provider.labels).toEqual(
      expect.arrayContaining([
        "Action required",
        "Plan for change",
        "Major change",
        "Admin impact",
      ]),
    );
    expect(items.find((item) => item.ref === "MC200002")?.class).toBe("message.retirement");
    // Prism365's own guess from the title is labeled as its own, not as Microsoft's.
    const inferred = items.find((item) => item.ref === "MC200003");
    expect(inferred?.provider.labels).not.toContain("Retirement");
    expect(inferred?.localLabels).toContain("Possible retirement (from the title)");
  });

  it("retires a Message center item once Microsoft stops listing it", async () => {
    await setup();
    await poll();
    // Microsoft still lists other posts, but not this one.
    messages = [
      message({
        id: "MC200050",
        title: "Another post",
        tags: ["Retirement"],
        actionRequiredByDateTime: null,
      }),
    ];
    await poll();
    const item = (await attention("service.message")).find((entry) => entry.ref === "MC200001");
    expect(item?.provider.active).toBe(false);
    expect(item?.changeNote).toBe("No longer listed by Microsoft.");
  });

  it("skips a module whose permission is not granted and says so", async () => {
    await setup(["ServiceHealth.Read.All"]);
    await poll();
    expect(await attention("service.message")).toEqual([]);
    expect(await attention("service.incident")).toHaveLength(1);
  });
});

describe("service health collector under odd and partial data", () => {
  it("stores odd Microsoft records instead of failing the run on them", async () => {
    issues = [
      issue({
        id: "EX100009",
        title: "Odd\u0000 record",
        service: `Exchange Online ${"x".repeat(150)}`,
        impactDescription: "i".repeat(1_900),
        posts: [
          {
            createdDateTime: "2026-10-22T11:30:00Z",
            postType: "regular",
            description: {
              contentType: "html",
              content: "<p>Workaround: try &#x110000; later</p>",
            },
          },
        ],
      }),
      issue({ id: "EX100010" }),
    ];
    messages = [message({ tags: ["", "Admin impact"], services: [""] })];
    await setup();
    await poll();
    expect(
      (await database.tenant(tenantId, (sql) => latestRuns(sql, tenantId))).map(
        (run) => run.status,
      ),
    ).not.toContain("failed");
    expect((await attention("service.incident")).map((item) => item.ref).sort()).toEqual([
      "EX100009",
      "EX100010",
    ]);
    expect(await attention("service.message")).toHaveLength(1);
  });

  it("closes an incident Microsoft no longer lists as open, even after a long gap", async () => {
    await setup();
    await poll();
    // Resolved while collection was stopped, and outside the 7-day lookback by now.
    issues = [
      issue({ id: "EX100001", isResolved: true, lastModifiedDateTime: "2026-10-01T00:00:00Z" }),
    ];
    clock = new Date(start.getTime() + 20 * 86_400_000);
    await poll();
    const [item] = await attention("service.incident");
    expect(item?.provider.active).toBe(false);
    expect(item?.changeNote).toBe("Microsoft no longer lists this as open.");
  });

  it("does not close Message center items on an empty or partly unreadable read", async () => {
    await setup();
    await poll();
    messages = [];
    await poll();
    expect((await attention("service.message"))[0]?.provider.active).toBe(true);
    messages = [message(), { id: "", title: "" }];
    await poll();
    expect((await attention("service.message"))[0]?.provider.active).toBe(true);
  });

  it("says when Prism365's own rule, not Microsoft, stops selecting a post", async () => {
    messages = [message({ isMajorChange: false })];
    await setup();
    await poll();
    // Still listed by Microsoft, but its deadline is now long past and nothing else selects it.
    messages = [
      message({ isMajorChange: false }),
      message({
        id: "MC200099",
        title: "Unrelated recent post",
        tags: ["Retirement"],
        actionRequiredByDateTime: null,
      }),
    ];
    clock = new Date("2026-12-01T00:00:00Z");
    await poll();
    const item = (await attention("service.message")).find((entry) => entry.ref === "MC200001");
    expect(item?.provider.active).toBe(false);
    expect(item?.changeNote).toContain("No longer selected by Prism365");
  });
});

describe("service health check state", () => {
  const now = new Date("2026-10-22T12:00:00Z");
  const base = {
    enabled: true,
    granted: true,
    collecting: true,
    run: null,
    lastSuccess: null,
    staleAfterMs: 15 * 60_000,
  };
  const run = (status: "failed" | "succeeded", startedAt: string) => ({
    module: "service.health",
    collector: "health",
    status,
    credentialKind: "delegated" as const,
    startedAt,
    finishedAt: startedAt,
    fetched: 0,
    inserted: 0,
    findings: 0,
    skipped: 0,
    errorCategory: status === "failed" ? "timeout" : null,
  });
  it("tells missing permissions, failures, staleness and never-checked apart from a clean check", () => {
    expect(checkState({ ...base, enabled: false }, now).state).toBe("not_enabled");
    expect(checkState({ ...base, granted: false }, now).state).toBe("missing_permission");
    expect(checkState(base, now).state).toBe("never");
    expect(
      checkState(
        {
          ...base,
          lastSuccess: "2026-10-22T11:00:00Z",
          run: run("failed", "2026-10-22T11:55:00Z"),
        },
        now,
      ),
    ).toMatchObject({ state: "failed", detail: expect.stringContaining("timeout") });
    expect(checkState({ ...base, lastSuccess: "2026-10-22T11:00:00Z" }, now).state).toBe("stale");
    expect(checkState({ ...base, lastSuccess: "2026-10-22T11:58:00Z" }, now).state).toBe("ok");
    expect(
      checkState(
        {
          ...base,
          lastSuccess: "2026-10-22T11:58:00Z",
          run: { ...run("succeeded", "2026-10-22T11:58:00Z"), status: "partial" },
        },
        now,
      ).state,
    ).toBe("partial");
  });
});

describe("Microsoft post text", () => {
  it("reads labeled sections from HTML and never keeps markup", () => {
    const text = plainText(
      "<p>User impact: Users can&#39;t join <b>meetings</b>.</p><p>Workaround: Use the web client.<br>Or the phone.</p><script>alert(1)</script>",
    );
    expect(text).not.toContain("<");
    expect(text).not.toContain("alert");
    const sections = postSections(text);
    expect(sections.get("user impact")).toBe("Users can't join meetings.");
    expect(sections.get("workaround")).toBe("Use the web client.\nOr the phone.");
  });

  it("leaves no tag behind when tags are nested or split", () => {
    for (const html of [
      "<scr<b>ipt>alert(1)</scr<b>ipt>",
      "<<i>script>alert(1)<</i>/script>",
      "<<<b>>>text<<</b>>>",
      "<p<p>>text</p</p>>",
    ]) {
      expect(plainText(html)).not.toMatch(/<[a-z!/]/i);
    }
  });
});
