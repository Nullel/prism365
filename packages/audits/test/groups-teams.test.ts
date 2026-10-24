import { auditRowSchema, queryAuditRows } from "@prism365/contracts";
import { describe, expect, it } from "vitest";
import {
  auditDefinition,
  createGroupsTeamsAudit,
  GROUP_KINDS,
  GROUP_STATUSES,
  groupsTeamsAudit,
  MAX_GROUPS,
  MAX_OWNER_LISTS,
  OWNER_CONCURRENCY,
} from "../src/index.ts";
import {
  daysAgo,
  expectValidTable,
  fake,
  find,
  now,
  paged,
  type Route,
} from "./apps-devices-usage-support.ts";

// Synthetic groups and teams on a.example, with invented owners and a Teams activity report.
// The report path answers 302 with a link on the download host, as Microsoft does, and the
// download host serves the CSV.

const audit = groupsTeamsAudit;
const REPORT_HOST = "reports.office.com";
const REPORT_PATH = "/v1.0/reports/getTeamsTeamActivityDetail(period='D90')";
const granted = ["Group.Read.All", "Reports.Read.All", "ReportSettings.Read.All"];
const DAY = 86_400_000;

const gid = (n: number) => `00000000-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;
const inDays = (days: number) => new Date(now.getTime() + days * DAY).toISOString();
const dayOf = (days: number) => new Date(now.getTime() - days * DAY).toISOString().slice(0, 10);

type Json = Record<string, unknown>;

const group = (n: number, over: Json = {}): Json => ({
  id: gid(n),
  displayName: `Group ${n}`,
  groupTypes: [],
  mailEnabled: false,
  securityEnabled: true,
  resourceProvisioningOptions: [],
  visibility: null,
  expirationDateTime: null,
  renewedDateTime: null,
  isAssignableToRole: null,
  onPremisesSyncEnabled: null,
  createdDateTime: daysAgo(500),
  ...over,
});
const m365 = (n: number, over: Json = {}): Json =>
  group(n, {
    groupTypes: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    visibility: "Private",
    renewedDateTime: daysAgo(100),
    ...over,
  });
const team = (n: number, over: Json = {}): Json =>
  m365(n, { resourceProvisioningOptions: ["Team"], ...over });
const owner = (n: number, over: Json = {}): Json => ({
  "@odata.type": "#microsoft.graph.user",
  id: gid(9_000 + n),
  displayName: `Owner ${n}`,
  userPrincipalName: `owner${n}@a.example`,
  accountEnabled: true,
  ...over,
});

// The headers of Microsoft's v1.0 page, verbatim (checked 2026-10-24).
const HEADER =
  "Report Refresh Date,Team Name,Team Id,Team Type,Last Activity Date,Report Period,Active Users,Active Channels,Guests,Reactions,Meetings Organized,Post Messages,Reply Messages,Channel Messages,Urgent Messages,Mentions,Active Shared Channels,Active External Users".split(
    ",",
  );
const field = (value: string) =>
  /[",\r\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
// A report file as Microsoft most likely writes it: a byte-order mark, then CRLF lines.
function csv(rows: readonly Record<string, string>[], header: readonly string[] = HEADER): string {
  const lines = [
    header.join(","),
    ...rows.map((row) =>
      header
        .map((name) => field(name === "Report Refresh Date" ? dayOf(2) : (row[name] ?? "")))
        .join(","),
    ),
  ];
  return `﻿${lines.join("\r\n")}\r\n`;
}
interface Use {
  id: string;
  last?: number | null;
  users?: number;
  channels?: number;
  guests?: number;
  name?: string;
}
const use = (value: Use): Record<string, string> => ({
  "Team Name": value.name ?? "A team",
  "Team Id": value.id,
  "Team Type": "Private",
  "Last Activity Date": value.last === null ? "" : dayOf(value.last ?? 3),
  "Report Period": "90",
  "Active Users": String(value.users ?? 5),
  "Active Channels": String(value.channels ?? 2),
  Guests: String(value.guests ?? 0),
});

type Answer = () => Response;
interface World {
  groups?: readonly unknown[];
  pageSize?: number;
  // Owners by group ID: a list, or the response itself.
  owners?: Record<string, readonly unknown[] | Answer>;
  // The report file, or the response for the report's path; undefined serves an empty report.
  report?: string | Answer;
  settings?: Answer;
  download?: Answer;
  groupsResponse?: (url: URL) => Response | undefined;
}

function tenant(world: World = {}): Route {
  const served = new Map<string, Answer>();
  let issued = 0;
  return (url) => {
    if (url.hostname === REPORT_HOST) return served.get(url.pathname)?.();
    if (url.pathname === "/v1.0/groups")
      return world.groupsResponse?.(url) ?? paged(url, world.groups ?? [], world.pageSize ?? 999);
    const owned = /^\/v1\.0\/groups\/([^/]+)\/owners$/.exec(url.pathname)?.[1];
    if (owned !== undefined) {
      const answer = world.owners?.[owned.toLowerCase()] ?? [];
      return typeof answer === "function" ? answer() : paged(url, answer, 100);
    }
    if (url.pathname === "/v1.0/admin/reportSettings")
      return world.settings?.() ?? Response.json({ displayConcealedNames: false });
    if (url.pathname !== REPORT_PATH) return undefined;
    const body = world.report ?? csv([]);
    if (typeof body === "function") return body();
    issued += 1;
    const link = `/data/download/teams-${issued}`;
    served.set(
      link,
      world.download ??
        (() => new Response(body, { headers: { "Content-Type": "application/octet-stream" } })),
    );
    return new Response(null, {
      status: 302,
      headers: { Location: `https://${REPORT_HOST}${link}` },
    });
  };
}

const status = (code: number, headers: Record<string, string> = {}): Answer => {
  return () => new Response("{}", { status: code, headers });
};

interface RunOptions {
  granted?: string[];
  limits?: Parameters<typeof createGroupsTeamsAudit>[0];
}
const run = (route: Route, options: RunOptions = {}) => {
  const subject = options.limits ? createGroupsTeamsAudit(options.limits) : audit;
  const world = fake(subject, route, { granted: options.granted ?? granted });
  return subject.run(world.context).then((result) => ({ ...result, requests: world.requests }));
};
const runWith = (world: World, options?: RunOptions) => run(tenant(world), options);
const note = (result: { notes: string[] }) => result.notes.join("\n");
const tile = (result: { summary: { label: string; value: number }[] }, label: string) =>
  result.summary.find((item) => item.label === label)?.value;
const cells = (result: { rows: { key: string; cells: Record<string, unknown> }[] }, n: number) =>
  find(result.rows, gid(n)).cells;
const ownerRequests = (requests: URL[]) =>
  requests.filter((url) => url.pathname.endsWith("/owners"));

// Thirteen groups, one for every kind and every status, with their owners and report rows.
const ALPHA = 1;
const BETA = 2;
const GAMMA = 3;
const DELTA = 4;
const EPSILON = 5;
const ZETA = 6;
const ETA = 7;
const THETA = 8;
const sink: World = {
  groups: [
    team(ALPHA, { displayName: "Alpha team" }),
    team(BETA, { displayName: "Beta team" }),
    team(GAMMA, { displayName: "Gamma team", visibility: "Public" }),
    team(DELTA, { displayName: "Delta team" }),
    team(EPSILON, { displayName: "Epsilon team" }),
    team(ZETA, { displayName: "Zeta team", visibility: "Public" }),
    m365(ETA, { displayName: "Eta group", expirationDateTime: inDays(10) }),
    m365(THETA, { displayName: "Theta group" }),
    group(9, { displayName: "Iota", groupTypes: ["DynamicMembership"] }),
    group(10, { displayName: "Kappa", mailEnabled: true }),
    group(11, { displayName: "Lambda", isAssignableToRole: true }),
    group(12, { displayName: "Mu", mailEnabled: true, securityEnabled: false }),
    group(13, { displayName: "Nu", onPremisesSyncEnabled: true }),
  ],
  owners: {
    [gid(ALPHA)]: [owner(1), owner(2)],
    [gid(BETA)]: [owner(3)],
    [gid(GAMMA)]: [owner(4), owner(5)],
    [gid(DELTA)]: [],
    [gid(EPSILON)]: [owner(6, { accountEnabled: false })],
    [gid(ZETA)]: [owner(7), owner(8)],
    [gid(ETA)]: [owner(1), owner(2), owner(3)],
    [gid(THETA)]: [],
    [gid(11)]: [owner(1)],
  },
  report: csv([
    use({ id: gid(ALPHA), guests: 2 }),
    use({ id: gid(BETA), guests: 3 }),
    use({ id: gid(GAMMA), last: 140 }),
    use({ id: gid(DELTA) }),
    use({ id: gid(EPSILON) }),
    // A team that was deleted since: the report lists it and the directory does not.
    use({ id: gid(99) }),
  ]),
};

describe("Groups and Teams", () => {
  it("publishes a valid table", async () => {
    const result = await expectValidTable(audit, tenant(sink), { granted });
    expect(result.rows).toHaveLength(13);
  });

  it("needs the group pack, with the usage reports and report settings packs optional, unscoped, no handoff", () => {
    expect(auditDefinition(audit)).toMatchObject({
      id: "groups.teams",
      title: "Groups and Teams",
      permissions: ["Group.Read.All"],
      optionalPermissions: ["Reports.Read.All", "ReportSettings.Read.All"],
      scoped: false,
      handoff: false,
    });
    expect(audit.packs).toEqual(["directory.groups.read"]);
    expect(audit.optionalPacks).toEqual(["reports.usage.read", "reports.settings.read"]);
    const facets = auditDefinition(audit)
      .columns.filter((column) => column.facet)
      .map((column) => column.key);
    expect(facets).toEqual(["kind", "status", "visibility"]);
  });

  it("asks for the group list with the fields it uses and the largest page Microsoft allows", async () => {
    const result = await runWith(sink);
    const list = result.requests.find((url) => url.pathname === "/v1.0/groups");
    expect(list?.searchParams.get("$top")).toBe("999");
    expect((list?.searchParams.get("$select") ?? "").split(",").sort()).toEqual(
      [
        "id",
        "displayName",
        "groupTypes",
        "mailEnabled",
        "securityEnabled",
        "resourceProvisioningOptions",
        "visibility",
        "expirationDateTime",
        "renewedDateTime",
        "isAssignableToRole",
        "onPremisesSyncEnabled",
        "createdDateTime",
      ].sort(),
    );
  });

  it("reads more than one page of groups", async () => {
    const groups = Array.from({ length: 7 }, (_, n) => group(n + 1));
    const result = await runWith({ groups, pageSize: 3 });
    expect(result.rows).toHaveLength(7);
    expect(result.requests.filter((url) => url.pathname === "/v1.0/groups")).toHaveLength(3);
    expect(result.complete).toBe(true);
  });

  it("reads nothing about members, and no other path under a group", async () => {
    const result = await runWith(sink);
    const paths = result.requests
      .filter((url) => url.hostname === "graph.microsoft.com")
      .map((url) => url.pathname.replace(/[0-9a-f-]{36}/, "{id}"));
    expect(new Set(paths)).toEqual(
      new Set(["/v1.0/groups", "/v1.0/groups/{id}/owners", REPORT_PATH]),
    );
  });
});

describe("the kind of a group", () => {
  const kind = async (...records: Json[]) => {
    const result = await runWith({ groups: records });
    return records.map((record) => find(result.rows, String(record.id)).cells.kind);
  };

  it("takes the first that matches, in the order Team, Microsoft 365, dynamic, mail-enabled, security, distribution", async () => {
    expect(
      await kind(
        // A team is also a Microsoft 365 group, and may be dynamic.
        team(1, { groupTypes: ["Unified", "DynamicMembership"] }),
        m365(2),
        m365(3, { groupTypes: ["Unified", "DynamicMembership"] }),
        group(4, { groupTypes: ["DynamicMembership"], mailEnabled: true }),
        group(5, { mailEnabled: true, securityEnabled: true }),
        group(6, { securityEnabled: true }),
        group(7, { mailEnabled: true, securityEnabled: false }),
        group(8, { mailEnabled: false, securityEnabled: false }),
      ),
    ).toEqual([
      "Team",
      "Microsoft 365 group",
      "Microsoft 365 group",
      "Dynamic security group",
      "Mail-enabled security group",
      "Security group",
      "Distribution list",
      "Distribution list",
    ]);
  });

  it("ignores letter case in Microsoft's labels", async () => {
    expect(
      await kind(
        group(1, { resourceProvisioningOptions: ["TEAM"], groupTypes: ["unified"] }),
        group(2, { groupTypes: ["UNIFIED"] }),
        group(3, { groupTypes: ["dynamicmembership"] }),
      ),
    ).toEqual(["Team", "Microsoft 365 group", "Dynamic security group"]);
  });

  it("does not make a team of a group whose provisioning options lack Team", async () => {
    expect(await kind(m365(1, { resourceProvisioningOptions: [] }))).toEqual([
      "Microsoft 365 group",
    ]);
    expect(await kind(m365(1, { resourceProvisioningOptions: ["Other"] }))).toEqual([
      "Microsoft 365 group",
    ]);
  });

  it("has a name for every kind the table can show", async () => {
    const result = await runWith(sink);
    const seen = new Set(result.rows.map((row) => row.cells.kind));
    expect([...seen].sort()).toEqual([...GROUP_KINDS].sort());
  });
});

describe("which groups have their owners read", () => {
  it("reads owners for teams, Microsoft 365 groups and role-assignable groups only", async () => {
    const result = await runWith({
      groups: [
        team(1),
        m365(2),
        group(3, { isAssignableToRole: true }),
        group(4, { isAssignableToRole: true, groupTypes: ["DynamicMembership"] }),
        group(5),
        group(6, { groupTypes: ["DynamicMembership"] }),
        group(7, { mailEnabled: true, securityEnabled: true }),
        group(8, { mailEnabled: true, securityEnabled: false }),
        group(9, { isAssignableToRole: false }),
      ],
      owners: { [gid(1)]: [owner(1)], [gid(2)]: [owner(2)], [gid(3)]: [owner(3)], [gid(4)]: [] },
    });
    const asked = ownerRequests(result.requests)
      .map((url) => url.pathname.split("/")[3])
      .sort();
    expect(asked).toEqual([gid(1), gid(2), gid(3), gid(4)]);
    for (const n of [5, 6, 7, 8, 9])
      expect(cells(result, n)).toMatchObject({
        owners: null,
        enabledOwners: null,
        ownerNames: null,
        status: "OK",
      });
    expect(cells(result, 3).owners).toBe(1);
    // Not asking for the others is by design, not a gap in the run.
    expect(result.complete).toBe(true);
  });

  it("asks for the fields it needs and no member list", async () => {
    const result = await runWith({ groups: [team(1)], owners: { [gid(1)]: [owner(1)] } });
    const request = ownerRequests(result.requests)[0];
    expect(request?.searchParams.get("$select")).toBe(
      "id,displayName,userPrincipalName,accountEnabled",
    );
    expect(result.requests.some((url) => url.pathname.includes("/members"))).toBe(false);
  });

  it("asks teams first and then the oldest groups, and stops at the limit", async () => {
    const groups = [
      m365(1, { createdDateTime: daysAgo(10) }),
      m365(2, { createdDateTime: daysAgo(900) }),
      team(3, { createdDateTime: daysAgo(50) }),
      m365(4, { createdDateTime: daysAgo(400) }),
      team(5, { createdDateTime: daysAgo(700) }),
      group(6, { isAssignableToRole: true, createdDateTime: daysAgo(800) }),
      m365(7, { createdDateTime: null }),
    ];
    const asked = async (maxOwnerLists: number) =>
      (await runWith({ groups }, { limits: { maxOwnerLists } })).requests
        .filter((url) => url.pathname.endsWith("/owners"))
        .map((url) => Number.parseInt(url.pathname.split("/")[3]?.slice(-12) ?? "0", 16))
        .sort((a, b) => a - b);
    // Both teams, then the oldest of the rest: 2 (900 days), 6 (800), 4 (400), 1 (10), 7 (unknown).
    expect(await asked(1)).toEqual([5]);
    expect(await asked(2)).toEqual([3, 5]);
    expect(await asked(3)).toEqual([2, 3, 5]);
    expect(await asked(5)).toEqual([2, 3, 4, 5, 6]);
    expect(await asked(7)).toEqual([1, 2, 3, 4, 5, 6, 7]);
  });

  it("shows Owners not read for the rest, says so, and makes the run partial", async () => {
    const groups = [team(1), m365(2), m365(3), m365(4, { createdDateTime: daysAgo(900) })];
    const result = await runWith(
      { groups, owners: { [gid(1)]: [owner(1)], [gid(4)]: [owner(2)] } },
      { limits: { maxOwnerLists: 2 } },
    );
    expect(cells(result, 2)).toMatchObject({ status: "Owners not read", owners: null });
    expect(cells(result, 3)).toMatchObject({ status: "Owners not read", owners: null });
    expect(cells(result, 4).owners).toBe(1);
    expect(note(result)).toContain("Owners were read for 2 of 4 groups, teams first");
    expect(note(result)).toContain("2 groups show Owners not read");
    expect(result.complete).toBe(false);
    expect(
      queryAuditRows(auditDefinition(audit), result.rows, { view: "owners-not-read" }),
    ).toHaveLength(2);
  });

  it("does not call a group with unread owners ownerless", async () => {
    const result = await runWith({ groups: [m365(1)] }, { limits: { maxOwnerLists: 0 } });
    expect(cells(result, 1).status).toBe("Owners not read");
    expect(tile(result, "Groups and teams with no owners")).toBe(0);
  });

  it("reads 3,000 owner lists at most, and no more than four at a time", async () => {
    expect(MAX_OWNER_LISTS).toBe(3_000);
    expect(OWNER_CONCURRENCY).toBe(4);
    let flying = 0;
    let peak = 0;
    const slow: Answer = () =>
      new Response(
        new ReadableStream({
          async pull(controller) {
            flying += 1;
            peak = Math.max(peak, flying);
            await new Promise((resolve) => setTimeout(resolve, 3));
            controller.enqueue(new TextEncoder().encode(JSON.stringify({ value: [owner(1)] })));
            controller.close();
            flying -= 1;
          },
        }),
        { headers: { "Content-Type": "application/json" } },
      );
    const groups = Array.from({ length: 24 }, (_, n) => m365(n + 1));
    const owners = Object.fromEntries(groups.map((record) => [String(record.id), slow]));
    const result = await runWith({ groups, owners });
    expect(result.rows).toHaveLength(24);
    expect(peak).toBe(OWNER_CONCURRENCY);
  });

  it("stops at 3,000 owner lists with the real limit", async () => {
    const groups = Array.from({ length: MAX_OWNER_LISTS + 1 }, (_, n) =>
      m365(n + 1, { createdDateTime: daysAgo(n + 1) }),
    );
    const result = await runWith({ groups });
    expect(ownerRequests(result.requests)).toHaveLength(MAX_OWNER_LISTS);
    expect(result.complete).toBe(false);
    expect(note(result)).toContain("Owners were read for 3,000 of 3,001 groups");
    // The newest group is the one left out, the oldest the first asked.
    expect(cells(result, 1).status).toBe("Owners not read");
    expect(cells(result, MAX_OWNER_LISTS + 1).status).toBe("No owners");
  }, 60_000);
});

describe("the owners of a group", () => {
  it("counts owners and enabled owners, and names the first five in order", async () => {
    const owners = [6, 2, 5, 1, 4, 3, 7].map((n) => owner(n));
    owners.push(owner(8, { accountEnabled: false }));
    const result = await runWith({ groups: [team(1)], owners: { [gid(1)]: owners } });
    expect(cells(result, 1)).toMatchObject({
      owners: 8,
      enabledOwners: 7,
      ownerNames: ["Owner 1", "Owner 2", "Owner 3", "Owner 4", "Owner 5"],
    });
  });

  it("names an owner by sign-in name when it has no display name, and leaves out one with neither", async () => {
    const result = await runWith({
      groups: [team(1)],
      owners: {
        [gid(1)]: [
          owner(1, { displayName: null }),
          owner(2, { displayName: null, userPrincipalName: null }),
        ],
      },
    });
    expect(cells(result, 1)).toMatchObject({ owners: 2, ownerNames: ["owner1@a.example"] });
  });

  it("does not count an owner whose state is not given as enabled, or as disabled", async () => {
    const result = await runWith({
      groups: [team(1), team(2), m365(3)],
      owners: {
        // An app owner carries no accountEnabled.
        [gid(1)]: [owner(1, { accountEnabled: undefined })],
        [gid(2)]: [owner(2, { accountEnabled: false }), owner(3, { accountEnabled: undefined })],
        [gid(3)]: [owner(4, { accountEnabled: false }), owner(5)],
      },
    });
    expect(cells(result, 1)).toMatchObject({ owners: 1, enabledOwners: 0, status: "One owner" });
    // One disabled and one unknown: not "only disabled".
    expect(cells(result, 2)).toMatchObject({
      owners: 2,
      enabledOwners: 0,
      status: "No usage data",
    });
    expect(cells(result, 3).status).toBe("OK");
  });

  it("calls a group whose every owner is disabled Only disabled owners", async () => {
    const result = await runWith({
      groups: [team(1), m365(2), m365(3)],
      owners: {
        [gid(1)]: [owner(1, { accountEnabled: false }), owner(2, { accountEnabled: false })],
        [gid(2)]: [owner(3, { accountEnabled: false })],
        [gid(3)]: [owner(4, { accountEnabled: false }), owner(5)],
      },
    });
    expect(cells(result, 1)).toMatchObject({ status: "Only disabled owners", enabledOwners: 0 });
    expect(cells(result, 2).status).toBe("Only disabled owners");
    expect(cells(result, 3).status).toBe("OK");
    expect(tile(result, "Groups with only disabled owners")).toBe(2);
  });

  it("calls a group with no owners No owners, only once the owners were read", async () => {
    const result = await runWith({
      groups: [team(1), m365(2), group(3, { isAssignableToRole: true })],
    });
    for (const n of [1, 2, 3])
      expect(cells(result, n)).toMatchObject({ status: "No owners", owners: 0, ownerNames: [] });
    expect(tile(result, "Groups and teams with no owners")).toBe(3);
  });

  it("calls only a team with one owner One owner", async () => {
    const result = await runWith({
      groups: [team(1), m365(2), group(3, { isAssignableToRole: true })],
      owners: { [gid(1)]: [owner(1)], [gid(2)]: [owner(2)], [gid(3)]: [owner(3)] },
    });
    expect(cells(result, 1).status).toBe("One owner");
    expect(cells(result, 2).status).toBe("OK");
    expect(cells(result, 3).status).toBe("OK");
    expect(tile(result, "Teams with one owner")).toBe(1);
  });

  it("makes the run partial when a list holds more owners than are read", async () => {
    const many = Array.from({ length: 250 }, (_, n) => owner(n + 1));
    const result = await runWith({ groups: [team(1)], owners: { [gid(1)]: many } });
    expect(cells(result, 1).owners).toBe(200);
    expect(note(result)).toContain("1 owner list holds more than Prism365 reads");
    expect(result.complete).toBe(false);
  });

  it("counts an owner record it cannot read, and makes the run partial", async () => {
    const result = await runWith({
      groups: [team(1)],
      owners: { [gid(1)]: [owner(1), { displayName: "no id" }, "oops"] },
    });
    expect(cells(result, 1).owners).toBe(1);
    expect(note(result)).toContain("could not read");
    expect(result.complete).toBe(false);
  });

  it("cleans names and bounds their length", async () => {
    const long = "x".repeat(900);
    const result = await runWith({
      groups: [team(1, { displayName: `Alpha\u0000 ${long}` })],
      owners: { [gid(1)]: [owner(1, { displayName: `Ada\u0000 ${long}` })] },
    });
    const row = cells(result, 1);
    expect(String(row.group)).not.toContain("\u0000");
    expect(String(row.group).length).toBeLessThanOrEqual(500);
    const names = row.ownerNames as string[];
    expect(names[0]).not.toContain("\u0000");
    expect(names[0]?.length).toBeLessThanOrEqual(200);
    for (const item of result.rows) auditRowSchema.parse(item);
  });
});

describe("when owners cannot be read", () => {
  const groups = Array.from({ length: 12 }, (_, n) =>
    m365(n + 1, { createdDateTime: daysAgo(n + 1) }),
  );

  it("stops asking and keeps what it has when Microsoft asks it to slow down", async () => {
    // The oldest five are asked first and answer; the rest are asked to wait.
    const first = groups.slice(-5).map((record) => String(record.id));
    const owners: Record<string, Answer> = {};
    for (const record of groups.slice(0, -5))
      owners[String(record.id)] = status(429, { "Retry-After": "600" });
    for (const id of first) owners[id] = () => Response.json({ value: [owner(1)] });
    const result = await runWith({ groups, owners });
    expect(note(result)).toContain(
      "Microsoft asked Prism365 to slow down, so owners stopped after",
    );
    expect(result.complete).toBe(false);
    expect(result.rows).toHaveLength(12);
    expect(cells(result, 12).owners).toBe(1);
    // Once Microsoft says to slow down, only the requests already in flight are left to finish.
    const asked = ownerRequests(result.requests).length;
    expect(asked).toBeLessThanOrEqual(first.length + OWNER_CONCURRENCY);
    expect(asked).toBeLessThan(groups.length);
  });

  it("stops at the first refusal, naming the status", async () => {
    const owners: Record<string, Answer> = {};
    for (const record of groups) owners[String(record.id)] = status(403);
    const result = await runWith({ groups, owners });
    expect(note(result)).toContain(
      "Microsoft refused to list owners (HTTP 403), so none were read.",
    );
    expect(ownerRequests(result.requests).length).toBeLessThanOrEqual(OWNER_CONCURRENCY);
    expect(cells(result, 1).status).toBe("Owners not read");
    expect(result.complete).toBe(false);
  });

  it("skips a group that has gone since the list was read", async () => {
    const result = await runWith({
      groups: [team(1), team(2)],
      owners: { [gid(1)]: status(404), [gid(2)]: [owner(1), owner(2)] },
    });
    expect(cells(result, 1).status).toBe("Owners not read");
    expect(cells(result, 2).owners).toBe(2);
    expect(note(result)).toContain("1 group had no answer for owners");
    expect(result.complete).toBe(false);
  });

  it("lets a session that ended reach the runner", async () => {
    await expect(
      runWith({ groups: [team(1)], owners: { [gid(1)]: status(401) } }),
    ).rejects.toMatchObject({ status: 401 });
  });
});

describe("the Teams report", () => {
  const one = (...rows: Record<string, string>[]) => ({
    groups: [team(1)],
    owners: { [gid(1)]: [owner(1), owner(2)] },
    report: csv(rows),
  });

  it("joins a team by Team Id and shows the report's numbers, not its names", async () => {
    const result = await runWith(
      one(
        use({
          id: gid(1),
          name: "Another name entirely",
          users: 12,
          channels: 4,
          guests: 3,
          last: 2,
        }),
      ),
    );
    expect(cells(result, 1)).toMatchObject({
      group: "Group 1",
      status: "OK",
      lastActivity: `${dayOf(2)}T00:00:00.000Z`,
      activeUsers: 12,
      activeChannels: 4,
      guests: 3,
    });
  });

  it("matches headers and the Team Id without regard to letter case", async () => {
    const id = gid(0xabcdef);
    const file = csv([use({ id: id.toUpperCase(), users: 7 })]);
    const [header, ...rest] = file.split("\r\n");
    for (const spelling of [header?.toLowerCase(), header?.toUpperCase()]) {
      const result = await runWith({
        groups: [team(0xabcdef)],
        owners: { [id]: [owner(1), owner(2)] },
        report: [spelling, ...rest].join("\r\n"),
      });
      expect(find(result.rows, id).cells).toMatchObject({ activeUsers: 7, status: "OK" });
    }
  });

  it("does not join by name: a team with the same name and another Team Id gets nothing", async () => {
    const result = await runWith(one(use({ id: gid(2), name: "Group 1", users: 9 })));
    expect(cells(result, 1)).toMatchObject({ activeUsers: null, status: "No usage data" });
  });

  it("says No usage data for a team with no report row, and nothing worse", async () => {
    const result = await runWith(one());
    expect(cells(result, 1)).toMatchObject({
      lastActivity: null,
      activeUsers: null,
      guests: null,
      activeChannels: null,
      status: "No usage data",
    });
    expect(tile(result, "Inactive teams")).toBe(0);
    expect(queryAuditRows(auditDefinition(audit), result.rows, { view: "no-usage" })).toHaveLength(
      1,
    );
    // The report was read; a team it does not list is not a gap in the run.
    expect(result.complete).toBe(true);
  });

  it("calls a team inactive when its row has no activity date or one over 90 days old", async () => {
    const days = async (last: number | null) =>
      cells(await runWith(one(use({ id: gid(1), last }))), 1).status;
    expect(await days(null)).toBe("Inactive 90 days");
    expect(await days(91)).toBe("Inactive 90 days");
    expect(await days(400)).toBe("Inactive 90 days");
    expect(await days(90)).toBe("OK");
    expect(await days(89)).toBe("OK");
    expect(await days(0)).toBe("OK");
  });

  it("never calls a team inactive on a date it cannot read, and says so", async () => {
    const bad = { ...use({ id: gid(1) }), "Last Activity Date": "sometime" };
    const result = await runWith(one(bad));
    expect(cells(result, 1)).toMatchObject({ status: "OK", lastActivity: null });
    expect(note(result)).toContain("1 team has a last activity date Prism365 could not read");
    expect(result.complete).toBe(false);
  });

  it("does not call a group that is not a team inactive", async () => {
    const result = await runWith({
      groups: [m365(1)],
      owners: { [gid(1)]: [owner(1)] },
      report: csv([use({ id: gid(1), last: 300 })]),
    });
    expect(cells(result, 1)).toMatchObject({ status: "OK", lastActivity: null });
    expect(note(result)).toContain("1 Microsoft 365 group appears in the Teams report");
  });

  it("takes the first row for a team named twice", async () => {
    const result = await runWith(
      one(use({ id: gid(1), users: 1 }), use({ id: gid(1), users: 99 })),
    );
    expect(cells(result, 1).activeUsers).toBe(1);
  });

  it("reads a report that lacks the optional columns", async () => {
    const header = ["Team Id", "Last Activity Date"];
    const result = await runWith({
      ...one(),
      report: csv([use({ id: gid(1) })], header),
    });
    expect(cells(result, 1)).toMatchObject({ status: "OK", activeUsers: null, guests: null });
    expect(result.complete).toBe(true);
  });

  it("reads the usage report only with the usage pack, and the privacy setting only when needed", async () => {
    const result = await runWith(one(use({ id: gid(1) })));
    expect(result.requests.filter((url) => url.pathname === REPORT_PATH)).toHaveLength(1);
    expect(result.requests.some((url) => url.pathname.endsWith("/reportSettings"))).toBe(false);
  });

  it("does not ask for the report without Reports.Read.All, and says every team has no usage data", async () => {
    const result = await runWith(one(use({ id: gid(1) })), { granted: ["Group.Read.All"] });
    expect(result.requests.some((url) => url.pathname.includes("/reports/"))).toBe(false);
    expect(cells(result, 1).status).toBe("No usage data");
    expect(note(result)).toContain("The Teams activity report was not read");
    expect(result.summary.map((item) => item.label)).toContain("Inactive teams (usage not read)");
    expect(result.summary.map((item) => item.label)).toContain(
      "Teams with guests (usage not read)",
    );
    // An optional pack that is off is a choice, not a failure.
    expect(result.complete).toBe(true);
  });

  it("carries on, says why, and is partial when the report cannot be read", async () => {
    const result = await runWith({ ...one(), report: status(403) });
    expect(note(result)).toContain(
      "The Teams activity report could not be read: Microsoft refused (HTTP 403).",
    );
    expect(cells(result, 1).status).toBe("No usage data");
    expect(result.complete).toBe(false);
  });

  it("carries on when the download link fails", async () => {
    const result = await runWith({ ...one(use({ id: gid(1) })), download: status(410) });
    expect(note(result)).toContain("could not be read: the download link failed");
    expect(result.rows).toHaveLength(1);
  });

  it("carries on when the report has no Team Id column", async () => {
    const header = HEADER.filter((name) => name !== "Team Id");
    const result = await runWith({ ...one(), report: csv([use({ id: gid(1) })], header) });
    expect(note(result)).toContain('it has no "Team Id" column');
    expect(result.complete).toBe(false);
  });

  const hashed = (n: number) => n.toString(16).padStart(32, "0");

  it("does not take a report with one odd team ID among real ones for a concealed one", async () => {
    const result = await runWith(one(use({ id: hashed(1) }), use({ id: gid(1), users: 7 })));
    expect(cells(result, 1)).toMatchObject({ activeUsers: 7, status: "OK" });
    expect(note(result)).not.toContain("hashed");
    expect(note(result)).toContain("1 row in the Teams report could not be read");
  });

  it("matches no team and unmasks nothing when the Team Ids are hashes", async () => {
    const result = await runWith({
      ...one(),
      report: csv([use({ id: hashed(1), name: hashed(2) }), use({ id: hashed(3) })]),
    });
    expect(cells(result, 1).status).toBe("No usage data");
    expect(note(result)).toContain(
      "Microsoft's privacy setting says names are shown, but the Teams report still returns hashed team identifiers",
    );
    expect(note(result)).toContain("turn off concealment");
    expect(result.complete).toBe(false);
    const cold = await runWith(
      { ...one(), report: csv([use({ id: hashed(1) })]) },
      { granted: ["Group.Read.All", "Reports.Read.All"] },
    );
    expect(note(cold)).toContain("identifiers look like hashes, so they are treated as concealed");
    const said = await runWith({
      ...one(),
      report: csv([use({ id: hashed(1) })]),
      settings: () => Response.json({ displayConcealedNames: true }),
    });
    expect(note(said)).toContain("Microsoft's report privacy setting says names are concealed");
  });

  it("counts rows it could not read", async () => {
    const good = csv([use({ id: gid(1) })]);
    const result = await runWith({
      ...one(),
      report: `${good}${csv([use({ id: "" })])
        .split("\r\n")
        .slice(1)
        .join("\r\n")}only,three,fields\r\n`,
    });
    expect(cells(result, 1).status).toBe("OK");
    expect(note(result)).toContain(
      "2 rows in the Teams report could not be read and are left out.",
    );
    expect(result.complete).toBe(false);
  });

  it("counts a number it cannot read and shows it empty", async () => {
    const result = await runWith(
      one({ ...use({ id: gid(1) }), "Active Users": "many", Guests: "-1" }),
    );
    expect(cells(result, 1)).toMatchObject({ activeUsers: null, guests: null });
    expect(note(result)).toContain("2 values in the Teams report could not be read");
    expect(result.complete).toBe(false);
  });
});

describe("expiry", () => {
  const expires = async (value: unknown) =>
    cells(
      await runWith({
        groups: [m365(1, { expirationDateTime: value })],
        owners: { [gid(1)]: [owner(1)] },
      }),
      1,
    );

  it("flags a group that expires within 30 days, to the millisecond, and one already past", async () => {
    expect((await expires(inDays(30))).status).toBe("Expires within 30 days");
    expect((await expires(new Date(now.getTime() + 30 * DAY + 1).toISOString())).status).toBe("OK");
    expect((await expires(inDays(31))).status).toBe("OK");
    expect((await expires(inDays(1))).status).toBe("Expires within 30 days");
    expect((await expires(inDays(-5))).status).toBe("Expires within 30 days");
    expect((await expires(null)).status).toBe("OK");
  });

  it("shows the dates and counts the tile", async () => {
    const result = await runWith({
      groups: [
        m365(1, { expirationDateTime: inDays(12), renewedDateTime: daysAgo(350) }),
        m365(2, { expirationDateTime: inDays(200) }),
      ],
      owners: { [gid(1)]: [owner(1)], [gid(2)]: [owner(1)] },
    });
    expect(cells(result, 1)).toMatchObject({ expires: inDays(12), renewed: daysAgo(350) });
    expect(tile(result, "Groups expiring in 30 days")).toBe(1);
  });

  it("does not call a group expiring on a date it cannot read, and says so", async () => {
    const result = await runWith({
      groups: [m365(1, { expirationDateTime: "soon" })],
      owners: { [gid(1)]: [owner(1)] },
    });
    expect(cells(result, 1)).toMatchObject({ status: "OK", expires: null });
    expect(note(result)).toContain("1 date on groups could not be read");
    expect(result.complete).toBe(false);
  });

  it("puts no owners, only disabled owners and one owner before inactive, and those before expiry", async () => {
    const stale = csv([
      use({ id: gid(1), last: 200 }),
      use({ id: gid(2), last: 200 }),
      use({ id: gid(3), last: 200 }),
    ]);
    const soon = { expirationDateTime: inDays(3) };
    const result = await runWith({
      groups: [team(1, soon), team(2, soon), team(3, soon), m365(4, soon)],
      owners: {
        [gid(1)]: [],
        [gid(2)]: [owner(1, { accountEnabled: false })],
        [gid(3)]: [owner(2), owner(3)],
        [gid(4)]: [owner(4)],
      },
      report: stale,
    });
    expect(cells(result, 1).status).toBe("No owners");
    expect(cells(result, 2).status).toBe("Only disabled owners");
    // Inactive comes before expiring; both tags stay.
    expect(cells(result, 3).status).toBe("Inactive 90 days");
    expect(cells(result, 4).status).toBe("Expires within 30 days");
    const tags = find(result.rows, gid(1)).tags;
    expect(tags).toEqual(expect.arrayContaining(["no-owners", "inactive", "expiring"]));
  });

  it("puts one owner before inactive", async () => {
    const result = await runWith({
      groups: [team(1)],
      owners: { [gid(1)]: [owner(1)] },
      report: csv([use({ id: gid(1), last: 200 })]),
    });
    expect(cells(result, 1).status).toBe("One owner");
  });
});

describe("the columns and tiles", () => {
  it("shows the dates, visibility, role-assignable and synced flags", async () => {
    const result = await runWith({
      groups: [
        team(1, { visibility: "Public", createdDateTime: daysAgo(30) }),
        m365(2, { visibility: "HiddenMembership" }),
        group(3, { visibility: "Weird", isAssignableToRole: true, onPremisesSyncEnabled: true }),
        group(4, { onPremisesSyncEnabled: false, isAssignableToRole: false }),
        group(5, { displayName: null, createdDateTime: null, isAssignableToRole: undefined }),
      ],
      owners: { [gid(1)]: [owner(1), owner(2)], [gid(2)]: [owner(1)], [gid(3)]: [owner(1)] },
    });
    expect(cells(result, 1)).toMatchObject({
      visibility: "Public",
      created: daysAgo(30),
      synced: false,
    });
    expect(cells(result, 2).visibility).toBe("Hidden membership");
    expect(cells(result, 3)).toMatchObject({
      visibility: "Weird",
      roleAssignable: true,
      synced: true,
    });
    expect(cells(result, 4)).toMatchObject({
      visibility: null,
      roleAssignable: false,
      synced: false,
    });
    expect(cells(result, 5)).toMatchObject({
      group: "Unnamed group",
      created: null,
      roleAssignable: null,
    });
  });

  it("counts every tile and links each to a view that lists the same groups", async () => {
    const result = await runWith(sink);
    const definition = auditDefinition(audit);
    const expected: [string, number, string][] = [
      ["Groups and teams with no owners", 2, "no-owners"],
      ["Groups with only disabled owners", 1, "disabled-owners"],
      ["Teams with one owner", 2, "one-owner"],
      ["Teams with guests", 2, "guests"],
      ["Inactive teams", 1, "inactive"],
      ["Public teams", 2, "public-teams"],
      ["Groups expiring in 30 days", 1, "expiring"],
      ["Role-assignable groups", 1, "role-assignable"],
    ];
    expect(result.summary.map((item) => item.label)).toEqual(expected.map(([label]) => label));
    for (const [label, value, view] of expected) {
      expect(tile(result, label), label).toBe(value);
      expect(result.summary.find((item) => item.label === label)?.view).toBe(view);
      expect(queryAuditRows(definition, result.rows, { view }), view).toHaveLength(value);
    }
    expect(result.complete).toBe(true);
  });

  it("gives the sink one group for each status", async () => {
    const result = await runWith(sink);
    const statuses = new Set(result.rows.map((row) => row.cells.status));
    expect([...statuses].sort()).toEqual(
      GROUP_STATUSES.filter((item) => item !== "Owners not read").sort(),
    );
    expect(cells(result, ALPHA).status).toBe("OK");
    expect(cells(result, BETA).status).toBe("One owner");
    expect(cells(result, GAMMA).status).toBe("Inactive 90 days");
    expect(cells(result, DELTA).status).toBe("No owners");
    expect(cells(result, EPSILON).status).toBe("Only disabled owners");
    expect(cells(result, ZETA).status).toBe("No usage data");
    expect(cells(result, ETA).status).toBe("Expires within 30 days");
    expect(cells(result, THETA).status).toBe("No owners");
  });

  it("lists the worst first and by name within a status", async () => {
    const result = await runWith(sink);
    expect(result.rows.slice(0, 2).map((row) => row.cells.group)).toEqual([
      "Delta team",
      "Theta group",
    ]);
    expect(result.rows.at(-1)?.cells.status).toBe("OK");
  });

  it("filters by kind, status and visibility", async () => {
    const result = await runWith(sink);
    const definition = auditDefinition(audit);
    const keys = (facets: Record<string, string>) =>
      queryAuditRows(definition, result.rows, { view: "all", facets }).map(
        (row) => row.cells.group,
      );
    expect(keys({ kind: "Distribution list" })).toEqual(["Mu"]);
    expect(keys({ status: "Only disabled owners" })).toEqual(["Epsilon team"]);
    expect(keys({ visibility: "Public" }).sort()).toEqual(["Gamma team", "Zeta team"]);
  });

  it("counts only teams among the public groups", async () => {
    const result = await runWith({
      groups: [
        team(1, { visibility: "Public" }),
        m365(2, { visibility: "Public" }),
        group(3, { visibility: "Public" }),
        team(4, { visibility: "Private" }),
      ],
    });
    expect(tile(result, "Public teams")).toBe(1);
    expect(
      queryAuditRows(auditDefinition(audit), result.rows, { view: "public-teams" }).map(
        (row) => row.key,
      ),
    ).toEqual([gid(1)]);
  });

  it("gives a group no page of its own", async () => {
    const result = await runWith(sink);
    expect(result.rows.every((row) => row.personId === null)).toBe(true);
  });
});

describe("limits and damaged records", () => {
  it("stops at 20,000 groups and says the rest was not read", async () => {
    expect(MAX_GROUPS).toBe(20_000);
    const groups = Array.from({ length: MAX_GROUPS + 1 }, (_, n) => group(n + 1));
    const result = await runWith({ groups });
    expect(result.rows).toHaveLength(MAX_GROUPS);
    expect(result.complete).toBe(false);
    expect(note(result)).toContain(
      "Stopped at 20,000 groups; the rest of the directory was not read.",
    );
  }, 60_000);

  it("reads exactly 20,000 groups whole", async () => {
    const groups = Array.from({ length: MAX_GROUPS }, (_, n) => group(n + 1));
    const result = await runWith({ groups });
    expect(result.rows).toHaveLength(MAX_GROUPS);
    expect(result.complete).toBe(true);
    expect(note(result)).not.toContain("Stopped at");
  }, 60_000);

  it("keeps what it read when Microsoft asks it to slow down between pages", async () => {
    const groups = Array.from({ length: 6 }, (_, n) => group(n + 1));
    const result = await runWith({
      groups,
      pageSize: 3,
      groupsResponse: (url) =>
        url.searchParams.has("$skiptoken")
          ? new Response("{}", { status: 429, headers: { "Retry-After": "600" } })
          : undefined,
    });
    expect(result.rows).toHaveLength(3);
    expect(note(result)).toContain("Microsoft asked Prism365 to slow down after 3 groups");
    expect(result.complete).toBe(false);
  });

  it("leaves out a group it cannot read, and a repeated one, and says so", async () => {
    const result = await runWith({
      groups: [
        group(1),
        { ...group(2), id: "not-a-guid" },
        { displayName: "no id" },
        "oops",
        group(1, { displayName: "Again" }),
      ],
    });
    expect(result.rows.map((row) => row.key)).toEqual([gid(1)]);
    expect(cells(result, 1).group).toBe("Group 1");
    expect(note(result)).toContain("3 groups could not be read and are left out.");
    expect(result.complete).toBe(false);
  });

  it("states what it does not read", async () => {
    const result = await runWith(sink);
    expect(note(result)).toContain("Members are not read");
    expect(note(result)).toContain(
      "Owners are read for Microsoft 365 groups, teams and groups that can be given a directory role",
    );
  });
});
