import { type AuditRow, cleanText, personIdSchema } from "@prism365/contracts";
import { z } from "zod";
import type { Audit } from "./definition.ts";
import { DAY_MS } from "./directory.ts";
import { firstAccepted } from "./graph-pages.ts";

// Devices enrolled in Intune: who has them, whether they pass compliance, when they last checked
// in. Microsoft Graph v1.0 documentation verified 2026-10-22:
// https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice?view=graph-rest-1.0

const deviceContext = {
  feature: "devices.intune.read",
  permission: "DeviceManagementManagedDevices.Read.All",
} as const;

// The client reads at most 10,000 records per collection.
export const MAX_DEVICES = 10_000;
export const STALE_DAYS = 30;

const text = z
  .string()
  .transform((value) => cleanText(value.slice(0, 500)))
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
const flag = z
  .boolean()
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
// Microsoft documents jailBroken as a string ("Unknown", "False", "True"). A boolean is accepted
// too, in case a tenant is answered by a newer shape.
const flagText = z
  .union([z.string(), z.boolean()])
  .nullish()
  .catch(null)
  .transform((value) => (value === null || value === undefined ? null : cleanText(String(value))));
const time = z
  .string()
  .datetime({ offset: true })
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);

const deviceSchema = z.object({
  id: z
    .string()
    .min(1)
    .max(200)
    .regex(/^[A-Za-z0-9._@:-]+$/),
  deviceName: text,
  userId: text,
  userPrincipalName: text,
  userDisplayName: text,
  operatingSystem: text,
  osVersion: text,
  complianceState: text,
  managedDeviceOwnerType: text,
  isEncrypted: flag,
  jailBroken: flagText,
  lastSyncDateTime: time,
  enrolledDateTime: time,
  wiFiMacAddress: text,
  serialNumber: text,
  model: text,
  manufacturer: text,
});
type Device = z.output<typeof deviceSchema>;

// Only the fields the audit uses. Microsoft's own documentation does not say that a list accepts
// $select, so a tenant that answers 400 is asked again without it.
const fields = [
  "id",
  "deviceName",
  "userId",
  "userPrincipalName",
  "userDisplayName",
  "operatingSystem",
  "osVersion",
  "complianceState",
  "managedDeviceOwnerType",
  "isEncrypted",
  "jailBroken",
  "lastSyncDateTime",
  "enrolledDateTime",
  "wiFiMacAddress",
  "serialNumber",
  "model",
  "manufacturer",
];
const forms = [
  `deviceManagement/managedDevices?$select=${fields.join(",")}&$top=500`,
  "deviceManagement/managedDevices?$top=200",
  "deviceManagement/managedDevices",
];

const complianceLabels: Record<string, string> = {
  compliant: "Compliant",
  noncompliant: "Not compliant",
  ingraceperiod: "In grace period",
  conflict: "Conflict",
  error: "Error",
  configmanager: "Managed by Configuration Manager",
  unknown: "Unknown",
};

export function complianceLabel(state: string | null): string {
  return complianceLabels[(state ?? "").toLowerCase()] ?? "Unknown";
}

export function ownershipLabel(type: string | null): "Company" | "Personal" | "Unknown" {
  const value = type?.toLowerCase();
  if (value === "company") return "Company";
  if (value === "personal") return "Personal";
  return "Unknown";
}

// Microsoft sends "True", "False" or "Unknown".
export function jailbrokenLabel(value: string | null): "Yes" | "No" | "Unknown" {
  const state = value?.trim().toLowerCase();
  if (state === "true") return "Yes";
  if (state === "false") return "No";
  return "Unknown";
}

// Intune writes 0001-01-01 for a device that has never checked in, and nothing a device could have
// sent before 2000 is a real date.
const EARLIEST = Date.parse("2000-01-01T00:00:00Z");

export interface CheckIn {
  at: string | null;
  days: number | null;
  // Microsoft has no check-in on record for the device.
  never: boolean;
}

// When a device last contacted Intune. A date that cannot be real, such as one in the future,
// is left unknown rather than counted.
export function checkInOf(value: string | null, now: Date): CheckIn {
  if (value === null) return { at: null, days: null, never: false };
  const at = Date.parse(value);
  if (Number.isNaN(at) || at > now.getTime() + DAY_MS)
    return { at: null, days: null, never: false };
  if (at < EARLIEST) return { at: null, days: null, never: true };
  return {
    at: new Date(at).toISOString(),
    days: Math.max(0, Math.floor((now.getTime() - at) / DAY_MS)),
    never: false,
  };
}

function enrolledOn(value: string | null): string | null {
  if (value === null) return null;
  const at = Date.parse(value);
  return Number.isNaN(at) || at < EARLIEST ? null : new Date(at).toISOString();
}

// Intune gives a device with no primary user the all-zero ID.
function ownerId(userId: string | null): string | null {
  if (!userId || /^0{8}-0{4}-0{4}-0{4}-0{12}$/.test(userId)) return null;
  return personIdSchema.safeParse(userId).success ? userId : null;
}

function rowOf(device: Device, now: Date): AuditRow {
  const compliance = complianceLabel(device.complianceState);
  const ownership = ownershipLabel(device.managedDeviceOwnerType);
  const jailbroken = jailbrokenLabel(device.jailBroken);
  const checkIn = checkInOf(device.lastSyncDateTime, now);
  const personId = ownerId(device.userId);
  const hasOwner = Boolean(device.userDisplayName ?? device.userPrincipalName);
  return {
    key: device.id,
    personId,
    tags: [
      ...(compliance === "Not compliant" ? ["not-compliant"] : []),
      ...(checkIn.never || (checkIn.days !== null && checkIn.days >= STALE_DAYS) ? ["stale"] : []),
      ...(ownership === "Personal" ? ["personal"] : []),
      ...(device.isEncrypted === false ? ["not-encrypted"] : []),
      ...(jailbroken === "Yes" ? ["jailbroken"] : []),
    ],
    cells: {
      device: device.deviceName,
      person: hasOwner
        ? { name: device.userDisplayName, principal: device.userPrincipalName }
        : { name: "No primary user", principal: null },
      os: device.operatingSystem ?? "Unknown",
      osVersion: device.osVersion,
      ownership,
      compliance,
      lastCheckIn: checkIn.at,
      days: checkIn.days,
      enrolled: enrolledOn(device.enrolledDateTime),
      encrypted: device.isEncrypted,
      jailbroken,
      wifiMac: device.wiFiMacAddress,
      serial: device.serialNumber,
      model: device.model,
      manufacturer: device.manufacturer,
    },
  };
}

export const managedDevicesAudit: Audit = {
  definition: {
    id: "devices.managed",
    version: 1,
    title: "Managed devices",
    description:
      "Devices enrolled in Intune: who has them, whether they pass compliance, when they last checked in, and which are personal, unencrypted or jailbroken.",
    category: "Devices",
    subject: { one: "device", many: "devices" },
    columns: [
      { key: "device", label: "Device", kind: "text", empty: "Unnamed device" },
      { key: "person", label: "Owner", kind: "person" },
      { key: "os", label: "Operating system", kind: "enum", facet: true },
      { key: "osVersion", label: "Version", kind: "text", empty: "Unknown" },
      {
        key: "compliance",
        label: "Compliance",
        kind: "enum",
        facet: true,
        help: "Intune's own verdict against the compliance policies assigned to the device. In grace period means the device has time to fix what failed.",
      },
      { key: "ownership", label: "Ownership", kind: "enum", facet: true },
      {
        key: "lastCheckIn",
        label: "Last check-in",
        kind: "date",
        empty: "None recorded",
        help: "When the device last contacted Intune. It says the device is reachable, not that anyone used it.",
      },
      {
        key: "days",
        label: "Since check-in",
        kind: "days",
        empty: "Unknown",
        help: "Days since the last check-in. Empty when Microsoft has no readable date.",
      },
      { key: "encrypted", label: "Encrypted", kind: "boolean", empty: "Unknown" },
      {
        key: "jailbroken",
        label: "Jailbroken or rooted",
        kind: "enum",
        facet: true,
        help: "Microsoft reports this as True, False or Unknown. Unknown is common for devices Intune cannot inspect.",
      },
      { key: "enrolled", label: "Enrolled", kind: "date", hidden: true, empty: "Unknown" },
      { key: "wifiMac", label: "Wi-Fi MAC address", kind: "text", hidden: true },
      { key: "serial", label: "Serial number", kind: "text", hidden: true },
      { key: "model", label: "Model", kind: "text", hidden: true },
      { key: "manufacturer", label: "Manufacturer", kind: "text", hidden: true },
    ],
    views: [
      { id: "not-compliant", label: "Not compliant", tags: ["not-compliant"] },
      { id: "stale", label: "Stale (no check-in for 30 days)", tags: ["stale"] },
      { id: "personal", label: "Personal devices", tags: ["personal"] },
      { id: "not-encrypted", label: "Not encrypted", tags: ["not-encrypted"] },
      { id: "jailbroken", label: "Jailbroken or rooted", tags: ["jailbroken"] },
      { id: "all", label: "All", tags: [] },
    ],
    defaultView: "not-compliant",
    defaultSort: { column: "days", direction: "desc" },
    requirements: [
      "An active Intune license in the tenant, which Microsoft requires for its device API.",
      "A role that can read Intune devices, such as Global Reader, Security Reader, Helpdesk Administrator or Intune Administrator, or the Read Only Operator role in Intune.",
    ],
  },
  packs: ["devices.intune.read"],
  optionalPacks: [],
  scoped: false,
  forbidden:
    "Microsoft refused to list managed devices. It needs DeviceManagementManagedDevices.Read.All, an Intune license in the tenant, and a role that can read Intune devices, such as Global Reader, Security Reader, Helpdesk Administrator or Intune Administrator.",
  async run(context) {
    context.signal.throwIfAborted();
    const result = await firstAccepted(forms, (path) =>
      context.graph.getCollection(path, deviceContext, {
        schema: z.unknown(),
        maxPages: 100,
        maxRecords: MAX_DEVICES,
      }),
    );
    let skipped = 0;
    const rows: AuditRow[] = [];
    for (const raw of result.items) {
      const parsed = deviceSchema.safeParse(raw);
      if (!parsed.success) {
        skipped += 1;
        continue;
      }
      rows.push(rowOf(parsed.data, context.now));
    }
    rows.sort((a, b) =>
      String(a.cells.device ?? "")
        .toLowerCase()
        .localeCompare(String(b.cells.device ?? "").toLowerCase()),
    );
    const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
    const unknownDates = rows.filter(
      (row) => row.cells.days === null && !row.tags.includes("stale"),
    ).length;
    const any = rows.length > 0;
    const tone = (value: number, level: "warning" | "critical") =>
      value > 0 ? level : any ? "good" : "neutral";
    return {
      rows,
      complete: result.complete && skipped === 0,
      summary: [
        {
          label: "Jailbroken or rooted",
          value: count("jailbroken"),
          tone: tone(count("jailbroken"), "critical"),
          view: "jailbroken",
        },
        {
          label: "Not compliant",
          value: count("not-compliant"),
          tone: tone(count("not-compliant"), "warning"),
          view: "not-compliant",
        },
        {
          label: "No check-in for 30 days",
          value: count("stale"),
          tone: tone(count("stale"), "warning"),
          view: "stale",
        },
        {
          label: "Not encrypted",
          value: count("not-encrypted"),
          tone: "neutral",
          view: "not-encrypted",
        },
        {
          label: "Personal devices",
          value: count("personal"),
          tone: "neutral",
          view: "personal",
        },
      ],
      notes: [
        ...(any
          ? []
          : [
              "Microsoft returned no managed devices. Either none are enrolled in Intune, or the tenant has no Intune license, which Microsoft's device API requires.",
            ]),
        "The Ethernet MAC address is not shown. Microsoft returns it only when a device is read on its own, and this audit does not make a request per device.",
        "Owner is the primary user Intune has on record. Shared and kiosk devices have none.",
        "Encrypted and jailbroken are what the device reported to Intune. Unknown means it did not say, which is common for some platforms.",
        ...(unknownDates > 0
          ? [
              `${unknownDates} ${unknownDates === 1 ? "device has" : "devices have"} no readable check-in date. ${unknownDates === 1 ? "It is" : "They are"} listed but never counted as stale.`,
            ]
          : []),
        ...(skipped > 0
          ? [
              `${skipped} ${skipped === 1 ? "device" : "devices"} could not be read and ${skipped === 1 ? "is" : "are"} left out.`,
            ]
          : []),
        ...(result.complete
          ? []
          : [`Stopped at ${MAX_DEVICES.toLocaleString("en-US")} devices; the rest were not read.`]),
      ],
    };
  },
};
