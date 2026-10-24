import { queryAuditRows } from "@prism365/contracts";
import { GraphRequestError } from "@prism365/graph";
import { describe, expect, it } from "vitest";
import {
  auditDefinition,
  checkInOf,
  complianceLabel,
  jailbrokenLabel,
  MAX_DEVICES,
  managedDevicesAudit,
  ownershipLabel,
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

// Synthetic devices: invented names, owners on example.invalid, made-up IDs and serial numbers.
const device = (id: string, extra: Record<string, unknown> = {}) => ({
  id,
  deviceName: `LAB-${id}`,
  userId: "11111111-0000-4000-8000-000000000001",
  userPrincipalName: "avery@example.invalid",
  userDisplayName: "Avery Active",
  operatingSystem: "Windows",
  osVersion: "10.0.26100",
  complianceState: "compliant",
  managedDeviceOwnerType: "company",
  isEncrypted: true,
  jailBroken: "False",
  lastSyncDateTime: daysAgo(1),
  enrolledDateTime: daysAgo(400),
  wiFiMacAddress: "02AABBCCDD01",
  serialNumber: `SN-${id}`,
  model: "Aerolite 14",
  manufacturer: "Kestrel Works",
  ethernetMacAddress: null,
  ...extra,
});

const devices = [
  device("ok"),
  device("bad", { complianceState: "noncompliant" }),
  device("grace", { complianceState: "inGracePeriod" }),
  device("config", { complianceState: "configManager" }),
  device("weird", { complianceState: "somethingNew" }),
  device("stale", { lastSyncDateTime: daysAgo(45) }),
  device("edge", { lastSyncDateTime: daysAgo(30) }),
  device("fresh-edge", { lastSyncDateTime: daysAgo(29) }),
  device("never", { lastSyncDateTime: "0001-01-01T00:00:00Z" }),
  device("future", { lastSyncDateTime: new Date(now.getTime() + 9 * 86_400_000).toISOString() }),
  device("personal", {
    managedDeviceOwnerType: "personal",
    isEncrypted: false,
    operatingSystem: "macOS",
    userId: "11111111-0000-4000-8000-000000000002",
    userPrincipalName: "sam@example.invalid",
    userDisplayName: "Sam Stale",
  }),
  device("jail", {
    operatingSystem: "iOS",
    jailBroken: "True",
    isEncrypted: null,
    complianceState: "noncompliant",
  }),
  device("unknown-jail", { jailBroken: "Unknown" }),
  device("kiosk", {
    userId: "00000000-0000-0000-0000-000000000000",
    userPrincipalName: null,
    userDisplayName: null,
    managedDeviceOwnerType: "unknown",
  }),
  // Unreadable: no id.
  { deviceName: "no id" },
];

const tenant: Route = (url) =>
  url.pathname === "/v1.0/deviceManagement/managedDevices" ? paged(url, devices, 500) : undefined;

const granted = ["DeviceManagementManagedDevices.Read.All"];
const run = (route: Route = tenant) =>
  managedDevicesAudit.run(fake(managedDevicesAudit, route, { granted }).context);

describe("managed devices", () => {
  it("publishes a valid table", async () => {
    await expectValidTable(managedDevicesAudit, tenant, { granted });
  });

  it("declares the Intune pack and nothing else", () => {
    expect(auditDefinition(managedDevicesAudit)).toMatchObject({
      id: "devices.managed",
      category: "Devices",
      permissions: ["DeviceManagementManagedDevices.Read.All"],
      optionalPermissions: [],
      scoped: false,
    });
  });

  it("lists each device with its owner, state and check-in", async () => {
    const result = await run();
    expect(result.rows).toHaveLength(devices.length - 1);
    expect(find(result.rows, "ok")).toMatchObject({
      personId: "11111111-0000-4000-8000-000000000001",
      tags: [],
      cells: {
        device: "LAB-ok",
        person: { name: "Avery Active", principal: "avery@example.invalid" },
        os: "Windows",
        osVersion: "10.0.26100",
        ownership: "Company",
        compliance: "Compliant",
        days: 1,
        encrypted: true,
        jailbroken: "No",
        wifiMac: "02AABBCCDD01",
        serial: "SN-ok",
        model: "Aerolite 14",
        manufacturer: "Kestrel Works",
      },
    });
    expect(find(result.rows, "ok").cells.lastCheckIn).toBe(daysAgo(1));
  });

  it("maps every compliance state Microsoft documents", () => {
    expect(
      [
        "compliant",
        "noncompliant",
        "inGracePeriod",
        "conflict",
        "error",
        "configManager",
        "unknown",
      ].map(complianceLabel),
    ).toEqual([
      "Compliant",
      "Not compliant",
      "In grace period",
      "Conflict",
      "Error",
      "Managed by Configuration Manager",
      "Unknown",
    ]);
    expect(complianceLabel("unknownFutureValue")).toBe("Unknown");
    expect(complianceLabel(null)).toBe("Unknown");
    expect(ownershipLabel("company")).toBe("Company");
    expect(ownershipLabel("Personal")).toBe("Personal");
    expect(ownershipLabel("unknown")).toBe("Unknown");
    expect(ownershipLabel(null)).toBe("Unknown");
  });

  it("maps Microsoft's jailBroken string, which is not a boolean", () => {
    expect(["True", "False", "Unknown", "true", " FALSE ", "", null].map(jailbrokenLabel)).toEqual([
      "Yes",
      "No",
      "Unknown",
      "Yes",
      "No",
      "Unknown",
      "Unknown",
    ]);
    expect(jailbrokenLabel("maybe")).toBe("Unknown");
  });

  it("flags only a device Microsoft says is not compliant", async () => {
    const result = await run();
    const flagged = result.rows.filter((row) => row.tags.includes("not-compliant"));
    expect(flagged.map((row) => row.key).sort()).toEqual(["bad", "jail"]);
    for (const key of ["grace", "config", "weird"])
      expect(find(result.rows, key).tags).not.toContain("not-compliant");
    expect(find(result.rows, "grace").cells.compliance).toBe("In grace period");
    expect(find(result.rows, "config").cells.compliance).toBe("Managed by Configuration Manager");
    expect(find(result.rows, "weird").cells.compliance).toBe("Unknown");
  });

  it("counts a device as stale from 30 days, and one that never checked in", async () => {
    const result = await run();
    const stale = result.rows.filter((row) => row.tags.includes("stale"));
    expect(stale.map((row) => row.key).sort()).toEqual(["edge", "never", "stale"]);
    expect(find(result.rows, "fresh-edge").tags).not.toContain("stale");
    expect(find(result.rows, "never").cells).toMatchObject({ lastCheckIn: null, days: null });
  });

  it("does not count a check-in date that cannot be real", async () => {
    const result = await run();
    expect(find(result.rows, "future")).toMatchObject({
      tags: [],
      cells: { lastCheckIn: null, days: null },
    });
    expect(result.notes.join(" ")).toContain("1 device has no readable check-in date");
    expect(checkInOf(null, now)).toEqual({ at: null, days: null, never: false });
    expect(checkInOf("not a date", now)).toEqual({ at: null, days: null, never: false });
    expect(checkInOf("0001-01-01T00:00:00Z", now)).toEqual({ at: null, days: null, never: true });
  });

  it("tags personal, unencrypted and jailbroken devices", async () => {
    const result = await run();
    expect(find(result.rows, "personal")).toMatchObject({
      tags: ["personal", "not-encrypted"],
      cells: { ownership: "Personal", encrypted: false },
    });
    expect(find(result.rows, "jail")).toMatchObject({
      tags: expect.arrayContaining(["jailbroken", "not-compliant"]),
      cells: { jailbroken: "Yes", encrypted: null },
    });
    // An unknown encryption state is not "not encrypted", and Unknown is not a jailbreak.
    expect(find(result.rows, "jail").tags).not.toContain("not-encrypted");
    expect(find(result.rows, "unknown-jail").tags).toEqual([]);
    expect(find(result.rows, "unknown-jail").cells.jailbroken).toBe("Unknown");
  });

  it("links a device to its owner, and says when there is none", async () => {
    const result = await run();
    expect(find(result.rows, "personal").personId).toBe("11111111-0000-4000-8000-000000000002");
    expect(find(result.rows, "kiosk")).toMatchObject({
      personId: null,
      cells: { person: { name: "No primary user", principal: null }, ownership: "Unknown" },
    });
  });

  it("opens each view on the devices that belong in it", async () => {
    const result = await run();
    const definition = auditDefinition(managedDevicesAudit);
    const keys = (view: string) =>
      queryAuditRows(definition, result.rows, { view }).map((row) => row.key);
    expect(keys("not-compliant")).toEqual(expect.arrayContaining(["bad", "jail"]));
    expect(keys("stale").sort()).toEqual(["edge", "never", "stale"]);
    expect(keys("personal")).toEqual(["personal"]);
    expect(keys("not-encrypted")).toEqual(["personal"]);
    expect(keys("jailbroken")).toEqual(["jail"]);
    expect(keys("all")).toHaveLength(devices.length - 1);
    expect(definition.defaultView).toBe("not-compliant");
  });

  it("filters by operating system, compliance and ownership", async () => {
    const result = await run();
    const definition = auditDefinition(managedDevicesAudit);
    const facets = definition.columns.filter((column) => column.facet).map((column) => column.key);
    expect(facets).toEqual(expect.arrayContaining(["os", "compliance", "ownership"]));
    expect(
      queryAuditRows(definition, result.rows, { view: "all", facets: { os: "macOS" } }).map(
        (row) => row.key,
      ),
    ).toEqual(["personal"]);
    expect(
      queryAuditRows(definition, result.rows, {
        view: "all",
        facets: { compliance: "In grace period" },
      }).map((row) => row.key),
    ).toEqual(["grace"]);
  });

  it("summarises with the tiles in the contract's order and severity", async () => {
    const result = await run();
    expect(result.summary).toEqual([
      { label: "Jailbroken or rooted", value: 1, tone: "critical", view: "jailbroken" },
      { label: "Not compliant", value: 2, tone: "warning", view: "not-compliant" },
      { label: "No check-in for 30 days", value: 3, tone: "warning", view: "stale" },
      { label: "Not encrypted", value: 1, tone: "neutral", view: "not-encrypted" },
      { label: "Personal devices", value: 1, tone: "neutral", view: "personal" },
    ]);
  });

  it("shows a clean tenant as good, not as a finding", async () => {
    const result = await run((url) =>
      url.pathname === "/v1.0/deviceManagement/managedDevices"
        ? paged(url, [device("a"), device("b")])
        : undefined,
    );
    expect(result.summary.map((tile) => tile.tone)).toEqual([
      "good",
      "good",
      "good",
      "neutral",
      "neutral",
    ]);
    expect(result.complete).toBe(true);
  });

  it("says so when Microsoft returns no devices, and does not call that all clear", async () => {
    const result = await run((url) =>
      url.pathname === "/v1.0/deviceManagement/managedDevices" ? paged(url, []) : undefined,
    );
    expect(result.rows).toEqual([]);
    expect(result.summary.every((tile) => tile.value === 0 && tile.tone === "neutral")).toBe(true);
    expect(result.notes.join(" ")).toContain("Microsoft returned no managed devices");
    expect(result.notes.join(" ")).toContain("Intune license");
  });

  it("says that the Ethernet MAC address is not read, and does not read it", async () => {
    const ctx = fake(managedDevicesAudit, tenant, { granted });
    const result = await managedDevicesAudit.run(ctx.context);
    expect(result.notes.join(" ")).toContain("Ethernet MAC address is not shown");
    // One list, no request per device.
    expect(
      ctx.requests.every((url) => url.pathname === "/v1.0/deviceManagement/managedDevices"),
    ).toBe(true);
    expect(ctx.requests.some((url) => url.searchParams.get("$select")?.includes("ethernet"))).toBe(
      false,
    );
    expect(Object.keys(find(result.rows, "ok").cells)).not.toContain("ethernetMac");
  });

  it("leaves out a record it cannot read and says the result is partial", async () => {
    const result = await run();
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("1 device could not be read and is left out");
    expect(result.rows.some((row) => row.cells.device === "no id")).toBe(false);
  });

  it("cleans provider text and keeps odd values from failing the run", async () => {
    const result = await run((url) =>
      url.pathname === "/v1.0/deviceManagement/managedDevices"
        ? paged(url, [
            device("odd", {
              deviceName: "Lab\u0000 laptop",
              osVersion: 17,
              isEncrypted: "yes",
              jailBroken: true,
              lastSyncDateTime: 12,
              serialNumber: "x".repeat(900),
            }),
          ])
        : undefined,
    );
    const odd = find(result.rows, "odd");
    expect(odd.cells).toMatchObject({
      device: "Lab laptop",
      osVersion: null,
      encrypted: null,
      jailbroken: "Yes",
      lastCheckIn: null,
    });
    expect(String(odd.cells.serial)).toHaveLength(500);
  });

  it("asks for a slim list first and falls back when Microsoft rejects the query", async () => {
    const strict: Route = (url) => {
      if (url.pathname !== "/v1.0/deviceManagement/managedDevices") return undefined;
      return url.searchParams.has("$select") || url.searchParams.get("$top") === "500"
        ? new Response("{}", { status: 400 })
        : paged(url, devices.slice(0, 3));
    };
    const ctx = fake(managedDevicesAudit, strict, { granted });
    const result = await managedDevicesAudit.run(ctx.context);
    expect(result.rows).toHaveLength(3);
    expect(ctx.requests).toHaveLength(2);
    expect(ctx.requests[0]?.searchParams.get("$select")).toContain("complianceState");
    expect(ctx.requests[0]?.searchParams.get("$top")).toBe("500");
    expect(ctx.requests[1]?.searchParams.has("$select")).toBe(false);
    expect(ctx.requests[1]?.searchParams.get("$top")).toBe("200");
  });

  it("gives up with Microsoft's error when every form is rejected", async () => {
    const ctx = fake(managedDevicesAudit, () => new Response("{}", { status: 400 }), { granted });
    await expect(managedDevicesAudit.run(ctx.context)).rejects.toMatchObject({ status: 400 });
    expect(ctx.requests).toHaveLength(3);
  });

  it("fails with Microsoft's refusal, which the page explains", async () => {
    const ctx = fake(managedDevicesAudit, () => new Response("{}", { status: 403 }), { granted });
    const failure = managedDevicesAudit.run(ctx.context);
    await expect(failure).rejects.toBeInstanceOf(GraphRequestError);
    await expect(failure).rejects.toMatchObject({ status: 403 });
    // A refusal is not retried in a plainer form.
    expect(ctx.requests).toHaveLength(1);
    const text = managedDevicesAudit.forbidden;
    expect(text).toContain("DeviceManagementManagedDevices.Read.All");
    expect(text).toContain("Intune license");
    expect(text).toContain("Global Reader");
  });

  it("stops at the record limit and says so", async () => {
    const many = Array.from({ length: MAX_DEVICES + 500 }, (_, index) => ({
      id: `d${index}`,
      deviceName: `LAB-${index}`,
      complianceState: "compliant",
    }));
    const result = await run((url) =>
      url.pathname === "/v1.0/deviceManagement/managedDevices" ? paged(url, many, 500) : undefined,
    );
    expect(result.rows).toHaveLength(MAX_DEVICES);
    expect(result.complete).toBe(false);
    expect(result.notes.join(" ")).toContain("Stopped at 10,000 devices");
  });

  it("reads every page of a large list", async () => {
    const many = Array.from({ length: 1_250 }, (_, index) => ({
      id: `d${index}`,
      deviceName: `LAB-${index}`,
    }));
    const ctx = fake(
      managedDevicesAudit,
      (url) =>
        url.pathname === "/v1.0/deviceManagement/managedDevices"
          ? paged(url, many, 500)
          : undefined,
      { granted },
    );
    const result = await managedDevicesAudit.run(ctx.context);
    expect(result.rows).toHaveLength(1_250);
    expect(result.complete).toBe(true);
    expect(ctx.requests).toHaveLength(3);
  });
});
