import { createHash } from "node:crypto";
import type { DirectoryUser, SyntheticTenant } from "../tenant.ts";
import { guid } from "../tenant.ts";
import type { DemoGraphExtension } from "./types.ts";

// Intune managed devices for the synthetic tenant: two or three per person, a few that fail
// compliance, some that have gone quiet, a jailbroken phone and a personal laptop with no
// encryption. Hardware makers and models are invented, like everything else here.

const DAY = 86_400_000;
const HOUR = 3_600_000;
const ZERO_ID = "00000000-0000-0000-0000-000000000000";

// A number in [0, 1) that depends only on the label, so the same tenant always gets the same
// devices.
function unit(label: string): number {
  return createHash("sha256").update(`devices:${label}`).digest().readUInt32LE(0) / 4_294_967_296;
}

type Kind = "laptop" | "phone" | "tablet";

interface Hardware {
  operatingSystem: string;
  osVersion: string;
  manufacturer: string;
  model: string;
}

const models = {
  windows: [
    { manufacturer: "Kestrel Works", model: "Aerolite 14" },
    { manufacturer: "Kestrel Works", model: "Aerolite 16" },
    { manufacturer: "Kestrel Works", model: "Aerolite Pro 15" },
  ],
  mac: [
    { manufacturer: "Orchard Devices", model: "Orchard Book 14" },
    { manufacturer: "Orchard Devices", model: "Orchard Book 16" },
  ],
  ios: [
    { manufacturer: "Orchard Devices", model: "Orchard Phone 15" },
    { manufacturer: "Orchard Devices", model: "Orchard Phone 15 Max" },
  ],
  android: [
    { manufacturer: "Tamarind Mobile", model: "Tamarind T9" },
    { manufacturer: "Tamarind Mobile", model: "Tamarind T9 Plus" },
  ],
  ipad: [{ manufacturer: "Orchard Devices", model: "Orchard Slate 11" }],
} as const;

const windowsVersion = "10.0.26100.4652";

function chooseModel<T>(label: string, items: readonly T[]): T {
  return items[Math.floor(unit(`${label}:model`) * items.length)] as T;
}

function hardwareFor(label: string, kind: Kind): Hardware & { family: string } {
  if (kind === "tablet")
    return {
      family: "ipad",
      operatingSystem: "iPadOS",
      osVersion: "18.6",
      ...chooseModel(label, models.ipad),
    };
  if (kind === "phone")
    return unit(`${label}:platform`) < 0.6
      ? {
          family: "ios",
          operatingSystem: "iOS",
          osVersion: "18.6.2",
          ...chooseModel(label, models.ios),
        }
      : {
          family: "android",
          operatingSystem: "Android",
          osVersion: "15",
          ...chooseModel(label, models.android),
        };
  return unit(`${label}:platform`) < 0.75
    ? {
        family: "windows",
        operatingSystem: "Windows",
        osVersion: windowsVersion,
        ...chooseModel(label, models.windows),
      }
    : {
        family: "mac",
        operatingSystem: "macOS",
        osVersion: "15.6.1",
        ...chooseModel(label, models.mac),
      };
}

// A Wi-Fi MAC address as Intune writes it, with no separators. The first byte is 02, the range
// reserved for locally administered addresses, so none can belong to real hardware.
function macFor(label: string): string {
  const digits = createHash("sha256").update(`mac:${label}`).digest("hex").slice(0, 10);
  return `02${digits.toUpperCase()}`;
}

function serialFor(label: string): string {
  const digits = createHash("sha256").update(`serial:${label}`).digest("hex").slice(0, 9);
  return `LK${digits.toUpperCase()}`;
}

const suffix = { laptop: "L", phone: "P", tablet: "T" } as const;

interface ManagedDevice {
  id: string;
  deviceName: string;
  userId: string;
  userPrincipalName: string | null;
  userDisplayName: string | null;
  emailAddress: string | null;
  operatingSystem: string;
  osVersion: string;
  complianceState: string;
  managedDeviceOwnerType: string;
  isEncrypted: boolean;
  jailBroken: string;
  lastSyncDateTime: string;
  enrolledDateTime: string;
  wiFiMacAddress: string;
  ethernetMacAddress: null;
  serialNumber: string;
  model: string;
  manufacturer: string;
  managementAgent: string;
  deviceEnrollmentType: string;
  azureADDeviceId: string;
  azureADRegistered: boolean;
}

// Whose devices carry the planted findings, counted among the people who have devices.
const planted = {
  jailbrokenPhone: 2,
  personalLaptop: 6,
  quietLaptop: 9,
  unencryptedAndroid: 11,
  quietTablet: 13,
};

const complianceFor = (roll: number) =>
  roll < 0.05
    ? "noncompliant"
    : roll < 0.08
      ? "inGracePeriod"
      : roll < 0.09
        ? "error"
        : roll < 0.1
          ? "conflict"
          : "compliant";

function devicesOf(tenant: SyntheticTenant): ManagedDevice[] {
  const now = tenant.now.getTime();
  const people = tenant.users.filter(
    (user: DirectoryUser) =>
      user.userType === "Member" && user.accountEnabled && user.assignedLicenses.length > 0,
  );
  const devices: ManagedDevice[] = [];
  people.forEach((user, index) => {
    const kinds: Kind[] = ["laptop", "phone"];
    // The tablet planted as quiet always exists; other people have one now and then.
    if (unit(`${user.id}:tablet`) < 0.4 || index === planted.quietTablet) kinds.push("tablet");
    for (const kind of kinds) {
      const label = `${user.id}:${kind}`;
      const make = hardwareFor(label, kind);
      let { operatingSystem, osVersion, model, manufacturer } = make;
      let compliance = complianceFor(unit(`${label}:compliance`));
      let ownership = kind === "phone" && unit(`${label}:byod`) < 0.2 ? "personal" : "company";
      let encrypted = true;
      let jailBroken = make.family === "ios" || make.family === "android" ? "False" : "Unknown";
      let syncedAgo = unit(`${label}:sync`) * 2.5 * DAY;

      if (index === planted.jailbrokenPhone && kind === "phone") {
        jailBroken = "True";
        compliance = "noncompliant";
        ({ operatingSystem, osVersion, model, manufacturer } = {
          operatingSystem: "Android",
          osVersion: "14",
          ...models.android[0],
        });
      }
      if (index === planted.personalLaptop && kind === "laptop") {
        ownership = "personal";
        encrypted = false;
        compliance = "noncompliant";
        ({ operatingSystem, osVersion, model, manufacturer } = {
          operatingSystem: "Windows",
          osVersion: windowsVersion,
          ...models.windows[0],
        });
      }
      if (index === planted.unencryptedAndroid && kind === "phone") {
        encrypted = false;
        jailBroken = "Unknown";
        ({ operatingSystem, osVersion, model, manufacturer } = {
          operatingSystem: "Android",
          osVersion: "13",
          ...models.android[1],
        });
      }
      if (index === planted.quietLaptop && kind === "laptop") syncedAgo = 47 * DAY + 3 * HOUR;
      if (index === planted.quietTablet && kind === "tablet") syncedAgo = 33 * DAY + 5 * HOUR;

      const surname = user.surname
        .replace(/[^A-Za-z]/g, "")
        .toUpperCase()
        .slice(0, 8);
      devices.push({
        id: guid(`managed-device:${label}`),
        deviceName: `LARK-${surname}-${suffix[kind]}${1 + (index % 3)}`,
        userId: user.id,
        userPrincipalName: user.userPrincipalName,
        userDisplayName: user.displayName,
        emailAddress: user.mail,
        operatingSystem,
        osVersion,
        complianceState: compliance,
        managedDeviceOwnerType: ownership,
        isEncrypted: encrypted,
        jailBroken,
        lastSyncDateTime: new Date(now - syncedAgo).toISOString(),
        enrolledDateTime: new Date(
          now - (60 + unit(`${label}:enrolled`) * 640) * DAY,
        ).toISOString(),
        wiFiMacAddress: macFor(label),
        // Microsoft returns null for this on the list, whatever the device has.
        ethernetMacAddress: null,
        serialNumber: serialFor(label),
        model,
        manufacturer,
        managementAgent: "mdm",
        deviceEnrollmentType: ownership === "personal" ? "userEnrollment" : "windowsAzureADJoin",
        azureADDeviceId: guid(`aad-device:${label}`),
        azureADRegistered: true,
      });
    }
  });
  // Shared laboratory machines have no primary user: Intune gives them the all-zero user ID.
  for (const number of [1, 2]) {
    const make = models.windows[1];
    devices.push({
      id: guid(`managed-device:kiosk:${number}`),
      deviceName: `LARK-KIOSK-0${number}`,
      userId: ZERO_ID,
      userPrincipalName: null,
      userDisplayName: null,
      emailAddress: null,
      operatingSystem: "Windows",
      osVersion: windowsVersion,
      complianceState: "compliant",
      managedDeviceOwnerType: "company",
      isEncrypted: true,
      jailBroken: "Unknown",
      lastSyncDateTime: new Date(now - unit(`kiosk:${number}:sync`) * DAY).toISOString(),
      enrolledDateTime: new Date(now - (200 + number * 30) * DAY).toISOString(),
      wiFiMacAddress: macFor(`kiosk:${number}`),
      ethernetMacAddress: null,
      serialNumber: serialFor(`kiosk:${number}`),
      model: make.model,
      manufacturer: make.manufacturer,
      managementAgent: "mdm",
      deviceEnrollmentType: "windowsAzureADJoin",
      azureADDeviceId: guid(`aad-device:kiosk:${number}`),
      azureADRegistered: true,
    });
  }
  return devices;
}

// Built once per tenant: the audit reads the list, and Microsoft would return the same devices
// to every request.
const built = new WeakMap<SyntheticTenant, ManagedDevice[]>();

// Intune managed devices.
export const managedDevices: DemoGraphExtension = (request, tools) => {
  if (request.path !== "deviceManagement/managedDevices") return null;
  let devices = built.get(request.tenant);
  if (!devices) {
    devices = devicesOf(request.tenant);
    built.set(request.tenant, devices);
  }
  return tools.page(request.url, devices, request.select);
};
