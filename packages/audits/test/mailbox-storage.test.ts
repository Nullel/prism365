import { auditRowSchema, queryAuditRows } from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import {
  AuditError,
  auditDefinition,
  MAILBOX_STATUSES,
  type MailboxStatus,
  mailboxStatus,
  mailboxStorageAudit,
} from "../src/index.ts";
import { expectValidTable, fake, find, now, type Route } from "./apps-devices-usage-support.ts";

// Synthetic mailboxes on a.example and b.example, in a report file as Microsoft serves it: the
// report path answers 302 with a link on the download host, and the download host serves the CSV.

const audit = mailboxStorageAudit;
const REPORT_HOST = "reports.office.com";
const REPORT_PATH = "/v1.0/reports/getMailboxUsageDetail(period='D90')";
const granted = ["Reports.Read.All", "ReportSettings.Read.All"];
const GIB = 1_073_741_824;

const columnsOf = (list: string) => list.split(",");
// The v1.0 page's header with Recipient Type, which Microsoft's current page adds (checked
// 2026-10-24).
const HEADER = columnsOf(
  "Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Created Date,Last Activity Date,Item Count,Storage Used (Byte),Issue Warning Quota (Byte),Prohibit Send Quota (Byte),Prohibit Send/Receive Quota (Byte),Deleted Item Count,Deleted Item Size (Byte),Deleted Item Quota (Byte),Has Archive,Recipient Type,Report Period",
);

const dayOf = (days: number) =>
  new Date(now.getTime() - days * 86_400_000).toISOString().slice(0, 10);
const field = (value: string) =>
  /[",\r\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
// A report file as Microsoft most likely writes it: a byte-order mark, then CRLF lines. A column
// a row does not name is empty.
function csv(
  rows: readonly Record<string, string>[],
  header: readonly string[] = HEADER,
  refreshed = dayOf(2),
): string {
  const lines = [
    header.join(","),
    ...rows.map((row) =>
      header
        .map((name) => field(name === "Report Refresh Date" ? refreshed : (row[name] ?? "")))
        .join(","),
    ),
  ];
  return `﻿${lines.join("\r\n")}\r\n`;
}

interface Box {
  upn: string;
  name?: string;
  used: number | string;
  // Quotas in bytes; the usual Exchange Online ones unless a test says otherwise.
  warning?: number | string;
  send?: number | string;
  sendReceive?: number | string;
  archive?: string;
  type?: string;
  deleted?: string;
  last?: number | null;
  items?: number;
  deletedBytes?: number;
}
// Bytes are whole numbers in the report, whatever fraction of a gigabyte a test asks for.
const bytes = (value: number | string): string =>
  typeof value === "number" ? String(Math.round(value)) : value;
const box = (value: Box): Record<string, string> => ({
  "User Principal Name": value.upn,
  "Display Name": value.name ?? value.upn.split("@")[0] ?? "",
  "Is Deleted": value.deleted ?? "False",
  "Last Activity Date": value.last === null ? "" : dayOf(value.last ?? 3),
  "Item Count": String(value.items ?? 1_000),
  "Storage Used (Byte)": bytes(value.used),
  "Issue Warning Quota (Byte)": bytes(value.warning ?? 49 * GIB),
  "Prohibit Send Quota (Byte)": bytes(value.send ?? 49.5 * GIB),
  "Prohibit Send/Receive Quota (Byte)": bytes(value.sendReceive ?? 50 * GIB),
  "Deleted Item Size (Byte)": String(value.deletedBytes ?? 0),
  "Has Archive": value.archive ?? "False",
  "Recipient Type": value.type ?? "UserMailbox",
});

// Small quotas make the boundaries exact: 90 percent of 1,000 is 900, 80 percent is 800.
const SMALL = { warning: 1_000, send: 1_100, sendReceive: 1_200 };

type Answer = () => Response;
interface World {
  report?: string | Answer;
  settings?: Answer;
  download?: Answer;
}

function tenant(world: World = {}): Route {
  const served = new Map<string, Answer>();
  let issued = 0;
  const file = csv([
    box({ upn: "ada@a.example", used: 3 * GIB }),
    box({ upn: "bo@b.example", used: 2 * GIB }),
  ]);
  return (url) => {
    if (url.hostname === REPORT_HOST) return served.get(url.pathname)?.();
    if (url.pathname === "/v1.0/admin/reportSettings")
      return world.settings?.() ?? Response.json({ displayConcealedNames: false });
    if (url.pathname !== REPORT_PATH) return undefined;
    const body = world.report ?? file;
    if (typeof body === "function") return body();
    issued += 1;
    const link = `/data/download/mailbox-${issued}`;
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

const run = (route: Route = tenant(), options: { granted?: string[]; scope?: string[] } = {}) => {
  const world = fake(audit, route, {
    granted: options.granted ?? granted,
    scope: options.scope ? new Set(options.scope) : null,
  });
  return audit.run(world.context).then((result) => ({ ...result, requests: world.requests }));
};
const runBoxes = (rows: readonly Record<string, string>[], options?: Parameters<typeof run>[1]) =>
  run(tenant({ report: csv(rows) }), options);
const note = (result: { notes: string[] }) => result.notes.join("\n");
const tile = (result: { summary: { label: string; value: number }[] }, label: string) =>
  result.summary.find((item) => item.label === label)?.value;
const statusOf = (rows: { key: string; cells: Record<string, unknown> }[], key: string) =>
  find(rows, key).cells.status;

describe("Mailbox storage", () => {
  it("publishes a valid table", async () => {
    const result = await expectValidTable(audit, tenant(), { granted });
    expect(result.rows).toHaveLength(2);
  });

  it("needs only the usage reports pack, with the report settings pack optional, scoped, no handoff", () => {
    const definition = auditDefinition(audit);
    expect(definition).toMatchObject({
      id: "mail.mailbox-storage",
      title: "Mailbox storage",
      permissions: ["Reports.Read.All"],
      optionalPermissions: ["ReportSettings.Read.All"],
      scoped: true,
      handoff: false,
    });
    expect(audit.packs).toEqual(["reports.usage.read"]);
    expect(audit.optionalPacks).toEqual(["reports.settings.read"]);
    expect(definition.columns.some((column) => column.key === "domain")).toBe(true);
    expect(definition.columns.filter((column) => column.facet).map((column) => column.key)).toEqual(
      ["recipientType", "status", "domain"],
    );
  });

  it("reads the D90 mailbox report once and asks for nothing else", async () => {
    const result = await run();
    const paths = result.requests
      .filter((url) => url.hostname === "graph.microsoft.com")
      .map((url) => url.pathname);
    expect(paths.filter((path) => path.includes("/reports/"))).toEqual([REPORT_PATH]);
    expect(paths.filter((path) => !path.includes("/reports/"))).toEqual([
      "/v1.0/admin/reportSettings",
    ]);
    expect(result.requests.filter((url) => url.hostname === REPORT_HOST)).toHaveLength(1);
  });

  it("gives every status but OK a view of its own, so no problem is listed only in a table", () => {
    const definition = auditDefinition(audit);
    const tags = definition.views.flatMap((view) => view.tags);
    expect(MAILBOX_STATUSES).toHaveLength(7);
    expect(tags.sort()).toEqual([
      "cannot-send",
      "cannot-send-receive",
      "near-full",
      "near-full-no-archive",
      "over-warning",
      "quota-unknown",
    ]);
  });
});

describe("the status of a mailbox", () => {
  const ladder = (archive: boolean | null) => (used: number) =>
    mailboxStatus(used, { warning: 1_000, send: 1_100, sendReceive: 1_200 }, archive);

  it("goes from worst to best, and reaching a quota counts as being over it", () => {
    const status = ladder(true);
    expect(status(1_200)).toBe("Cannot send or receive");
    expect(status(5_000)).toBe("Cannot send or receive");
    expect(status(1_199)).toBe("Cannot send");
    expect(status(1_100)).toBe("Cannot send");
    expect(status(1_099)).toBe("Over warning");
    expect(status(1_000)).toBe("Over warning");
    expect(status(999)).toBe("Near full");
    expect(status(900)).toBe("Near full");
    expect(status(899)).toBe("OK");
    expect(status(0)).toBe("OK");
  });

  it("calls a mailbox near full without an archive from 80 percent, only when it has none", () => {
    expect(ladder(false)(900)).toBe("Near full");
    expect(ladder(false)(899)).toBe("Near full, no archive");
    expect(ladder(false)(800)).toBe("Near full, no archive");
    expect(ladder(false)(799)).toBe("OK");
    // An archive that is there, or not known, is not called missing.
    expect(ladder(true)(850)).toBe("OK");
    expect(ladder(null)(850)).toBe("OK");
  });

  it("skips a quota the report does not give, and says Quota unknown without the warning quota", () => {
    const none = { warning: null, send: null, sendReceive: null };
    expect(mailboxStatus(10 * GIB, none, false)).toBe("Quota unknown");
    expect(mailboxStatus(10, { ...none, send: 100 }, false)).toBe("Quota unknown");
    // What is over a quota the report does give is still said.
    expect(mailboxStatus(150, { ...none, send: 100 }, false)).toBe("Cannot send");
    expect(mailboxStatus(150, { ...none, sendReceive: 100 }, false)).toBe("Cannot send or receive");
    expect(mailboxStatus(150, { ...none, warning: 100 }, false)).toBe("Over warning");
  });

  it("is the same through a whole run, on the boundaries", async () => {
    const cases: [number, string, MailboxStatus][] = [
      [1_200, "False", "Cannot send or receive"],
      [1_199, "False", "Cannot send"],
      [1_100, "False", "Cannot send"],
      [1_099, "False", "Over warning"],
      [1_000, "False", "Over warning"],
      [999, "False", "Near full"],
      [900, "False", "Near full"],
      [899, "False", "Near full, no archive"],
      [800, "False", "Near full, no archive"],
      [799, "False", "OK"],
      [899, "True", "OK"],
      [850, "", "OK"],
    ];
    const result = await runBoxes(
      cases.map(([used, archive], position) =>
        box({ upn: `m${position}@a.example`, used, archive, ...SMALL }),
      ),
    );
    for (const [position, [used, archive, expected]] of cases.entries())
      expect(statusOf(result.rows, `m${position}@a.example`), `${used} ${archive}`).toBe(expected);
  });

  it("reads blank quotas, and a quota of 0, as Quota unknown", async () => {
    const result = await runBoxes([
      box({ upn: "blank@a.example", used: 10 * GIB, warning: "", send: "", sendReceive: "" }),
      box({ upn: "zero@a.example", used: 10 * GIB, warning: 0, send: 0, sendReceive: 0 }),
      box({ upn: "nowarn@a.example", used: 10 * GIB, warning: "" }),
      box({
        upn: "overblank@a.example",
        used: 60 * GIB,
        warning: "",
        send: "",
        sendReceive: 50 * GIB,
      }),
    ]);
    expect(statusOf(result.rows, "blank@a.example")).toBe("Quota unknown");
    expect(statusOf(result.rows, "zero@a.example")).toBe("Quota unknown");
    expect(statusOf(result.rows, "nowarn@a.example")).toBe("Quota unknown");
    expect(statusOf(result.rows, "overblank@a.example")).toBe("Cannot send or receive");
    expect(find(result.rows, "blank@a.example").cells).toMatchObject({
      warningQuota: null,
      sendQuota: null,
      sendReceiveQuota: null,
      percent: null,
    });
    expect(tile(result, "Cannot send or receive")).toBe(1);
    expect(note(result)).toContain("3 mailboxes have no warning quota in the report");
  });
});

describe("the rows", () => {
  it("shows sizes in gigabytes to one decimal and sorts by bytes", async () => {
    const result = await runBoxes([
      box({ upn: "small@a.example", used: Math.round(0.4 * GIB) }),
      box({ upn: "tie-low@a.example", used: Math.round(3.126 * GIB), deletedBytes: 2 * GIB }),
      box({ upn: "tie-high@a.example", used: Math.round(3.139 * GIB) }),
      box({ upn: "big@a.example", used: Math.round(44.26 * GIB) }),
    ]);
    expect(result.rows.map((row) => row.key)).toEqual([
      "big@a.example",
      "tie-high@a.example",
      "tie-low@a.example",
      "small@a.example",
    ]);
    expect(find(result.rows, "big@a.example").cells).toMatchObject({
      used: 44.3,
      warningQuota: 49,
      sendQuota: 49.5,
      sendReceiveQuota: 50,
    });
    expect(find(result.rows, "tie-low@a.example").cells).toMatchObject({
      used: 3.1,
      deletedSize: 2,
    });
    expect(find(result.rows, "tie-high@a.example").cells.used).toBe(3.1);
  });

  it("shows the percent of the prohibit send quota, past 100 when the mailbox is over it", async () => {
    const result = await runBoxes([
      box({ upn: "half@a.example", used: 24.75 * GIB }),
      box({ upn: "over@a.example", used: 52 * GIB }),
    ]);
    expect(find(result.rows, "half@a.example").cells.percent).toBe(50);
    expect(find(result.rows, "over@a.example").cells.percent).toBe(105.1);
  });

  it("carries the type, items, archive, last activity and domain, and a name for people to read", async () => {
    const result = await runBoxes([
      box({
        upn: "Lab.Results@A.Example",
        name: "Lab Results",
        used: GIB,
        type: "SharedMailbox",
        archive: "True",
        items: 4_200,
        last: 5,
      }),
      box({ upn: "quiet@a.example", used: GIB, last: null, archive: "" }),
    ]);
    const shared = find(result.rows, "lab.results@a.example");
    expect(shared.cells).toMatchObject({
      mailbox: "Lab.Results@A.Example",
      displayName: "Lab Results",
      recipientType: "SharedMailbox",
      archive: true,
      items: 4_200,
      lastActivity: `${dayOf(5)}T00:00:00.000Z`,
      domain: "a.example",
    });
    expect(shared.personId).toBeNull();
    expect(find(result.rows, "quiet@a.example").cells).toMatchObject({
      lastActivity: null,
      archive: null,
    });
    for (const item of result.rows) auditRowSchema.parse(item);
  });

  it("filters by status and by recipient type", async () => {
    const result = await runBoxes([
      box({ upn: "full@a.example", used: 51 * GIB, type: "SharedMailbox" }),
      box({ upn: "fine@a.example", used: GIB, type: "UserMailbox" }),
      box({ upn: "room@a.example", used: GIB, type: "RoomMailbox" }),
    ]);
    const definition = auditDefinition(audit);
    expect(
      queryAuditRows(definition, result.rows, { view: "cannot-send-receive" }).map((r) => r.key),
    ).toEqual(["full@a.example"]);
    expect(
      queryAuditRows(definition, result.rows, {
        view: "all",
        facets: { recipientType: "RoomMailbox" },
      }).map((r) => r.key),
    ).toEqual(["room@a.example"]);
  });

  it("leaves out deleted mailboxes, in any spelling, and says how many", async () => {
    const result = await runBoxes([
      box({ upn: "live@a.example", used: GIB }),
      box({ upn: "gone@a.example", used: 60 * GIB, deleted: "True" }),
      box({ upn: "gone2@a.example", used: 60 * GIB, deleted: "TRUE" }),
    ]);
    expect(result.rows.map((row) => row.key)).toEqual(["live@a.example"]);
    expect(tile(result, "Cannot send or receive")).toBe(0);
    expect(note(result)).toContain("2 deleted mailboxes in the report are not listed.");
    expect(result.complete).toBe(true);
  });

  it("matches the report's headers without regard to letter case", async () => {
    const file = csv([box({ upn: "ada@a.example", used: 51 * GIB, type: "RoomMailbox" })]);
    const [header, ...rest] = file.split("\r\n");
    for (const spelling of [header?.toLowerCase(), header?.toUpperCase()]) {
      const result = await run(tenant({ report: [spelling, ...rest].join("\r\n") }));
      expect(find(result.rows, "ada@a.example").cells).toMatchObject({
        status: "Cannot send or receive",
        recipientType: "RoomMailbox",
        used: 51,
      });
    }
  });

  it("reads a report with no Recipient Type column, and says the types are unknown", async () => {
    const header = HEADER.filter((name) => name !== "Recipient Type");
    const result = await run(
      tenant({ report: csv([box({ upn: "ada@a.example", used: GIB })], header) }),
    );
    expect(find(result.rows, "ada@a.example").cells.recipientType).toBeNull();
    expect(note(result)).toContain("no Recipient Type column");
    // Missing is unknown, not a failure.
    expect(result.complete).toBe(true);
  });

  it("says when the report has no archive column, and calls no mailbox near full without one", async () => {
    const header = HEADER.filter((name) => name !== "Has Archive");
    const result = await run(
      tenant({ report: csv([box({ upn: "ada@a.example", used: 850, ...SMALL })], header) }),
    );
    expect(statusOf(result.rows, "ada@a.example")).toBe("OK");
    expect(find(result.rows, "ada@a.example").cells.archive).toBeNull();
    expect(note(result)).toContain("no archive column");
  });

  it("keeps the first of two rows for one mailbox, and counts names in any letter case as one", async () => {
    const result = await runBoxes([
      box({ upn: "ada@a.example", used: 51 * GIB }),
      box({ upn: "ADA@a.example", used: GIB }),
    ]);
    expect(result.rows).toHaveLength(1);
    expect(statusOf(result.rows, "ada@a.example")).toBe("Cannot send or receive");
  });

  it("cleans names and bounds their length", async () => {
    const long = "x".repeat(900);
    const result = await runBoxes([
      box({ upn: "ada@a.example", name: `Ada\u0000 ${long}`, used: GIB, type: `Odd${long}` }),
    ]);
    const cells = find(result.rows, "ada@a.example").cells;
    expect(String(cells.displayName)).not.toContain("\u0000");
    expect(String(cells.displayName).length).toBeLessThanOrEqual(200);
    expect(String(cells.recipientType).length).toBeLessThanOrEqual(40);
  });
});

describe("the tiles", () => {
  it("counts each status, and the storage used in gigabytes", async () => {
    const result = await runBoxes([
      box({ upn: "a1@a.example", used: 51 * GIB }),
      box({ upn: "a2@a.example", used: 50 * GIB }),
      box({ upn: "a3@a.example", used: 49.7 * GIB }),
      box({ upn: "a4@a.example", used: 49.2 * GIB }),
      box({ upn: "a5@a.example", used: 46 * GIB }),
      box({ upn: "a6@a.example", used: 45 * GIB, archive: "True" }),
      box({ upn: "a7@a.example", used: 41 * GIB, archive: "False" }),
      box({ upn: "a8@a.example", used: 41 * GIB, archive: "True" }),
      box({ upn: "a9@a.example", used: 10 * GIB }),
    ]);
    expect(tile(result, "Cannot send or receive")).toBe(2);
    expect(tile(result, "Cannot send")).toBe(1);
    expect(tile(result, "Over warning")).toBe(1);
    expect(tile(result, "Near full")).toBe(2);
    expect(tile(result, "Near full, no archive")).toBe(1);
    expect(tile(result, "Total storage used (GB)")).toBe(
      Math.round(51 + 50 + 49.7 + 49.2 + 46 + 45 + 41 + 41 + 10),
    );
    const definition = auditDefinition(audit);
    for (const item of result.summary)
      if (item.view) expect(definition.views.map((view) => view.id)).toContain(item.view);
    const views = (id: string) =>
      queryAuditRows(definition, result.rows, { view: id }).map((row) => row.key);
    expect(views("near-full-no-archive")).toEqual(["a7@a.example"]);
    expect(views("cannot-send-receive")).toEqual(["a1@a.example", "a2@a.example"]);
  });
});

describe("the audit scope", () => {
  const rows = [
    box({ upn: "ada@a.example", used: 51 * GIB }),
    box({ upn: "bo@b.example", used: 49.7 * GIB }),
    box({ upn: "cy@b.example", used: 10 * GIB }),
  ];

  it("lists only the mailboxes in the chosen domains, and counts only those in the tiles", async () => {
    const result = await runBoxes(rows, { scope: ["b.example"] });
    expect(result.rows.map((row) => row.key)).toEqual(["bo@b.example", "cy@b.example"]);
    expect(tile(result, "Cannot send or receive")).toBe(0);
    expect(tile(result, "Cannot send")).toBe(1);
    expect(tile(result, "Total storage used (GB)")).toBe(60);
    expect(note(result)).toContain("Limited to accounts in: b.example.");
    expect(note(result)).toContain("1 account in another domain was left out by the audit scope.");
  });

  it("lists every domain it read, whatever the scope", async () => {
    const result = await runBoxes(rows, { scope: ["b.example"] });
    expect(result.domains).toEqual([
      { domain: "a.example", accounts: 1 },
      { domain: "b.example", accounts: 2 },
    ]);
  });

  it("covers every domain with no scope or an empty one", async () => {
    expect((await runBoxes(rows)).rows).toHaveLength(3);
    expect((await runBoxes(rows, { scope: [] })).rows).toHaveLength(3);
  });

  it("scopes by the domain of the sign-in name in any letter case", async () => {
    const result = await runBoxes([box({ upn: "Ada@A.EXAMPLE", used: GIB })], {
      scope: ["a.example"],
    });
    expect(result.rows).toHaveLength(1);
  });
});

describe("when names are concealed", () => {
  const hashed = (n: number) => n.toString(16).padStart(32, "0");
  const concealed = [
    box({ upn: hashed(1), used: 51 * GIB }),
    box({ upn: hashed(2), used: 49.7 * GIB }),
    box({ upn: hashed(3), used: GIB }),
  ];

  it("shows totals and no rows when the identifiers are hashes, and links no tile to a view", async () => {
    const result = await runBoxes(concealed, { granted: ["Reports.Read.All"] });
    expect(result.rows).toEqual([]);
    expect(tile(result, "Cannot send or receive")).toBe(1);
    expect(tile(result, "Cannot send")).toBe(1);
    expect(result.summary.every((item) => item.view === null)).toBe(true);
    expect(result.domains).toBeUndefined();
    expect(note(result)).toContain(
      "could not read the report privacy setting, and the report's identifiers look like hashes",
    );
    expect(JSON.stringify(result)).not.toContain(hashed(1));
  });

  it("does not take a report with one odd name among readable ones for a concealed one", async () => {
    const result = await runBoxes([
      box({ upn: hashed(1), used: GIB }),
      box({ upn: "ada@a.example", used: GIB }),
    ]);
    expect(result.rows).toHaveLength(2);
    expect(note(result)).not.toContain("concealed");
  });

  it("says the scope could not be applied, instead of applying it to hashes", async () => {
    const result = await runBoxes(concealed, {
      scope: ["a.example"],
      granted: ["Reports.Read.All"],
    });
    expect(result.rows).toEqual([]);
    expect(note(result)).toContain(
      "The audit scope could not be applied: the report gives no names to sort by domain.",
    );
    expect(tile(result, "Cannot send or receive")).toBe(1);
  });

  it("trusts Microsoft's setting even when the names look readable", async () => {
    const result = await run(
      tenant({
        report: csv([box({ upn: "ada@a.example", used: 51 * GIB })]),
        settings: () => Response.json({ displayConcealedNames: true }),
      }),
    );
    expect(result.rows).toEqual([]);
    expect(note(result)).toContain("Names are concealed in this tenant's Microsoft usage reports");
    expect(tile(result, "Cannot send or receive")).toBe(1);
  });

  it("says so when the setting shows names but the report is hashed", async () => {
    const result = await run(
      tenant({
        report: csv(concealed),
        settings: () => Response.json({ displayConcealedNames: false }),
      }),
    );
    expect(note(result)).toContain("privacy setting says names are shown");
    expect(result.rows).toEqual([]);
  });

  it("names the setting it could not read when readable names come back", async () => {
    const result = await run(tenant(), { granted: ["Reports.Read.All"] });
    expect(result.rows).toHaveLength(2);
    expect(note(result)).toContain("(it needs ReportSettings.Read.All)");
  });
});

describe("what the report gets wrong", () => {
  it("fails with a plain message when a column the audit cannot do without is missing", async () => {
    const header = HEADER.filter((name) => name !== "Storage Used (Byte)");
    const error = await run(
      tenant({ report: csv([box({ upn: "ada@a.example", used: GIB })], header) }),
    ).catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(AuditError);
    expect((error as AuditError).code).toBe("report_columns");
    expect((error as AuditError).message).toContain('no "Storage Used (Byte)" column');
  });

  it("leaves out a mailbox with no readable size or name, and makes the run partial", async () => {
    const result = await runBoxes([
      box({ upn: "ada@a.example", used: GIB }),
      box({ upn: "blank@a.example", used: "" }),
      box({ upn: "bad@a.example", used: "12 GB" }),
      box({ upn: "", used: GIB }),
    ]);
    expect(result.rows.map((row) => row.key)).toEqual(["ada@a.example"]);
    expect(note(result)).toContain("3 mailboxes in the report have no readable name or size");
    expect(result.complete).toBe(false);
  });

  it("treats a quota it cannot read as missing, and says so", async () => {
    const result = await runBoxes([
      box({ upn: "ada@a.example", used: 10 * GIB, warning: "lots", send: "", sendReceive: "" }),
    ]);
    expect(statusOf(result.rows, "ada@a.example")).toBe("Quota unknown");
    expect(note(result)).toContain("1 value in the report could not be read");
    expect(result.complete).toBe(false);
  });

  it("makes the run partial for a date it cannot read", async () => {
    const result = await runBoxes([
      { ...box({ upn: "ada@a.example", used: GIB }), "Last Activity Date": "someday" },
    ]);
    expect(find(result.rows, "ada@a.example").cells.lastActivity).toBeNull();
    expect(result.complete).toBe(false);
  });

  it("counts rows the CSV reader dropped", async () => {
    const good = csv([box({ upn: "ada@a.example", used: GIB })]);
    const result = await run(tenant({ report: `${good}only,three,fields\r\n` }));
    expect(result.rows).toHaveLength(1);
    expect(note(result)).toContain("1 row in the report could not be read and is left out.");
    expect(result.complete).toBe(false);
  });

  it("reads an empty report as no mailboxes", async () => {
    const result = await runBoxes([]);
    expect(result.rows).toEqual([]);
    expect(result.complete).toBe(true);
    expect(tile(result, "Total storage used (GB)")).toBe(0);
  });

  it("states the date the report was refreshed", async () => {
    const result = await run();
    expect(note(result)).toContain(`refreshed ${dayOf(2)}`);
  });
});

describe("when Microsoft cannot give the report", () => {
  it("lets a refusal reach the runner, which explains it", async () => {
    const error = await run(tenant({ report: () => new Response("{}", { status: 403 }) })).catch(
      (caught: unknown) => caught,
    );
    expect(error).toBeInstanceOf(GraphRequestError);
    expect((error as GraphRequestError).status).toBe(403);
  });

  it("lets a download that failed reach the runner", async () => {
    const error = await run(
      tenant({ download: () => new Response("gone", { status: 410 }) }),
    ).catch((caught: unknown) => caught);
    expect(error).toBeDefined();
    expect(error).not.toBeInstanceOf(AuditError);
  });

  it("stops at once when the run is stopped", async () => {
    const controller = new AbortController();
    controller.abort(new Error("stopped"));
    const world = fake(audit, tenant(), { granted, signal: controller.signal });
    await expect(audit.run(world.context)).rejects.toThrow("stopped");
    expect(world.requests).toHaveLength(0);
  });
});
