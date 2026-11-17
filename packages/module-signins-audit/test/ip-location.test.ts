import { randomUUID } from "node:crypto";
import { type GraphClock, GraphHttpClient } from "@prism365/graph";
import {
  type IpLocation,
  type IpLocationProvider,
  IpLocationUnavailableError,
  latestRuns,
  listFindings,
  ModuleRegistry,
  ModuleScheduler,
} from "@prism365/modules";
import { PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createSigninsAuditModule,
  defaultSigninsAuditConfig,
  type SigninsAuditConfig,
} from "../src/index.ts";

// IPv6 sign-in locations from an IP location provider (docs/modules/signins-audit.md §13).
// Synthetic tenant and accounts; addresses are documentation ranges (192.0.2.0/24, 2001:db8::/32).

vi.setConfig({ testTimeout: 30_000 });

const tenantId = randomUUID();
const actorId = randomUUID();
const start = new Date("2026-11-09T12:00:00Z");

const chicago = {
  city: "Chicago",
  state: "Illinois",
  countryOrRegion: "US",
  geoCoordinates: { latitude: 41.88, longitude: -87.63 },
};
const amsterdam = {
  city: "Amsterdam",
  state: "Noord-Holland",
  countryOrRegion: "NL",
  geoCoordinates: { latitude: 52.37, longitude: 4.9 },
};
const seattle = {
  city: "Seattle",
  state: "Washington",
  countryOrRegion: "US",
  geoCoordinates: { latitude: 47.61, longitude: -122.33 },
};
const chicagoAnswer: IpLocation = {
  country: "US",
  state: "Illinois",
  city: "Chicago",
  lat: 41.88,
  lon: -87.63,
  accuracyKm: 20,
  network: null,
};

// The run clock, moved only by the fake provider, so the lookup budget means the same on any
// machine. Graph request timers never fire: no fake answer here is slow.
let time = 1_000;
const testClock: GraphClock = {
  now: () => time,
  setTimeout: () => 0,
  clearTimeout: () => {},
};

let clock: Date;
let records: Record<string, unknown>[];
let database: PGliteModuleDatabase;
let scheduler: ModuleScheduler;

function signIn(at: Date, overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    id: randomUUID(),
    createdDateTime: at.toISOString(),
    userId: "user-1",
    userPrincipalName: "synthetic.user@example.invalid",
    userDisplayName: "Synthetic User",
    appDisplayName: "Office 365 Exchange Online",
    ipAddress: "192.0.2.10",
    clientAppUsed: "Browser",
    conditionalAccessStatus: "success",
    isInteractive: true,
    riskLevelDuringSignIn: "hidden",
    status: { errorCode: 0, failureReason: null },
    deviceDetail: { deviceId: "device-1", operatingSystem: "Windows", browser: "Edge" },
    location: chicago,
    ...overrides,
  };
}

// An unfamiliar device on an IPv6 address Microsoft places in Amsterdam.
const ipv6SignIn = (at: Date, ip: string, overrides: Record<string, unknown> = {}) =>
  signIn(at, {
    ipAddress: ip,
    location: amsterdam,
    deviceDetail: { deviceId: "", operatingSystem: "Linux", browser: "Firefox" },
    ...overrides,
  });

async function fakeFetch(input: RequestInfo | URL): Promise<Response> {
  const url = new URL(String(input));
  const match = /createdDateTime ge (\S+) and createdDateTime (le|lt) (\S+)/.exec(
    url.searchParams.get("$filter") ?? "",
  );
  if (!match?.[1] || !match[3]) return new Response("{}", { status: 400 });
  const from = Date.parse(match[1]);
  const to = Date.parse(match[3]);
  const before = match[2] === "lt";
  const value = records
    .filter((record) => {
      const at = Date.parse(String(record.createdDateTime));
      return at <= clock.getTime() && at >= from && (before ? at < to : at <= to);
    })
    .sort((a, b) => Date.parse(String(b.createdDateTime)) - Date.parse(String(a.createdDateTime)));
  return Response.json({ value });
}

interface FakeProvider {
  value: IpLocationProvider;
  calls: string[][];
}

function fakeProvider(
  answer: (network: string) => IpLocation | null,
  options: { fail?: IpLocationUnavailableError; delayMs?: number } = {},
): FakeProvider {
  const calls: string[][] = [];
  return {
    calls,
    value: {
      id: "test-geo",
      name: "Synthetic lookup",
      attribution: "Synthetic attribution line.",
      locate: async (networks) => {
        calls.push([...networks]);
        time += options.delayMs ?? 50;
        if (options.fail) throw options.fail;
        return new Map(networks.map((network) => [network, answer(network)]));
      },
    },
  };
}

async function setup(provider?: FakeProvider, overrides: Partial<SigninsAuditConfig> = {}) {
  const registry = new ModuleRegistry().register(
    createSigninsAuditModule(
      { ...defaultSigninsAuditConfig, backfillDays: 3, ...overrides },
      { clock: testClock.now },
    ),
  );
  database = await PGliteModuleDatabase.open({ scopes: registry.migrationScopes() });
  scheduler = new ModuleScheduler({
    enabledPacks: ["audit.signins.read"],
    registry,
    database,
    now: () => clock,
    graphClient: (getAccessToken) =>
      new GraphHttpClient({
        getAccessToken,
        fetchImplementation: fakeFetch,
        sleep: async () => {},
        clock: testClock,
      }),
    ...(provider ? { ipLocation: provider.value } : {}),
  });
  scheduler.bind({
    kind: "delegated",
    tenantId,
    actorId,
    grantedPermissions: async () => ["User.Read", "AuditLog.Read.All"],
    accessToken: async () => "synthetic-token",
  });
}

async function poll() {
  await scheduler.tick();
  clock = new Date(clock.getTime() + 5 * 60_000);
}

const minutesAgo = (minutes: number) => new Date(clock.getTime() - minutes * 60_000);

const findings = () =>
  database.tenant(tenantId, (sql) => listFindings(sql, tenantId, { module: "signins.audit" }));
const limitationsOf = (finding: { limitations: unknown } | undefined) =>
  (finding?.limitations ?? []) as string[];
const fieldsOf = (finding: { evidence: unknown } | undefined) =>
  ((finding?.evidence ?? []) as { fields: Record<string, unknown> }[])[0]?.fields;

async function stored(ip: string) {
  const result = await database.tenant(tenantId, (sql) =>
    sql.query<Record<string, unknown>>(
      `SELECT country, state, city, lat, lon, ms_country, ms_state, ms_city, location_source,
              location_lookup, processed
         FROM mod_signins_audit_event WHERE ip = $1 ORDER BY created_at`,
      [ip],
    ),
  );
  return result.rows;
}

// Microsoft's own point, kept when the provider's location replaced it.
async function microsoftPoint(ip: string) {
  const result = await database.tenant(tenantId, (sql) =>
    sql.query<Record<string, unknown>>(
      "SELECT ms_lat, ms_lon FROM mod_signins_audit_event WHERE ip = $1 ORDER BY created_at",
      [ip],
    ),
  );
  return result.rows;
}

beforeEach(() => {
  clock = start;
  time = 1_000;
  records = [];
  // Three days of ordinary IPv4 history: one account, one office, one device.
  for (let hours = 70; hours >= 2; hours -= 4)
    records.push(signIn(new Date(start.getTime() - hours * 3_600_000)));
});

afterEach(async () => {
  await scheduler?.stop();
  await database?.close();
});

describe("IPv6 sign-in locations", () => {
  it("flags Microsoft's wrong IPv6 location when no provider is configured, and says why it may be wrong", async () => {
    await setup();
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2:aaaa:bbbb:cccc:dddd"));
    await poll();
    const [finding] = await findings();
    expect(finding?.ruleId).toBe("signins.unfamiliar_signin");
    expect(limitationsOf(finding).join(" ")).toContain(
      "its location is Microsoft's, which is often wrong for IPv6",
    );
    expect(await stored("2001:db8:1:2:aaaa:bbbb:cccc:dddd")).toEqual([
      expect.objectContaining({
        country: "NL",
        location_source: "microsoft",
        location_lookup: null,
      }),
    ]);
  });

  it("replaces Microsoft's IPv6 location with the provider's, keeps Microsoft's, and raises nothing false", async () => {
    const provider = fakeProvider(() => chicagoAnswer);
    await setup(provider);
    await poll();
    // History is IPv4 only: nothing was sent.
    expect(provider.calls).toEqual([]);
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2:aaaa:bbbb:cccc:dddd"));
    await poll();

    // Only the /64 left, never the full address.
    expect(provider.calls).toEqual([["2001:db8:1:2::/64"]]);
    expect(await stored("2001:db8:1:2:aaaa:bbbb:cccc:dddd")).toEqual([
      {
        country: "US",
        state: "Illinois",
        city: "Chicago",
        lat: 41.88,
        lon: -87.63,
        ms_country: "NL",
        ms_state: "Noord-Holland",
        ms_city: "Amsterdam",
        location_source: "test-geo",
        location_lookup: "found",
        processed: true,
      },
    ]);
    // Microsoft's point is kept too, so a wrong provider answer can still be investigated.
    expect(await microsoftPoint("2001:db8:1:2:aaaa:bbbb:cccc:dddd")).toEqual([
      { ms_lat: amsterdam.geoCoordinates.latitude, ms_lon: amsterdam.geoCoordinates.longitude },
    ]);
    // A new address on a new device in the account's own city: below the finding threshold.
    expect(await findings()).toEqual([]);
    expect(await stored("192.0.2.10")).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ location_source: "microsoft", location_lookup: null }),
      ]),
    );
  });

  it("explains a corrected location in the finding, with the provider's attribution", async () => {
    // Microsoft says Amsterdam; the provider says Berlin: a real new country, from the provider.
    const provider = fakeProvider(() => ({
      ...chicagoAnswer,
      country: "DE",
      state: "Land Berlin",
      city: "Berlin",
      lat: 52.52,
      lon: 13.4,
    }));
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77"));
    await poll();
    const [finding] = await findings();
    expect(fieldsOf(finding)).toMatchObject({
      country: "DE",
      city: "Berlin",
      locationSource: "test-geo",
      locationLookup: "found",
      microsoftLocation: "Amsterdam, Noord-Holland, NL",
    });
    expect(String(fieldsOf(finding)?.signals)).toBe(
      "new.ip, new.country, new.state, travel.impossible",
    );
    expect(limitationsOf(finding)).toEqual(
      expect.arrayContaining([
        expect.stringContaining("its location is from Synthetic lookup rather than Microsoft"),
        expect.stringContaining(
          "its region is compared only with regions Synthetic lookup has named for this account before",
        ),
        "Synthetic attribution line.",
      ]),
    );
    expect(limitationsOf(finding).every((line) => line.length <= 300)).toBe(true);
  });

  it("keeps Microsoft's location when a coarse answer agrees on the country, so a real foreign sign-in is still flagged", async () => {
    // Microsoft is right (Amsterdam); the provider only knows the country.
    const provider = fakeProvider(() => ({
      country: "NL",
      state: null,
      city: null,
      lat: 52.13,
      lon: 5.29,
      accuracyKm: 500,
      network: null,
    }));
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77"));
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({
        country: "NL",
        state: "Noord-Holland",
        city: "Amsterdam",
        lat: 52.37,
        lon: 4.9,
        ms_country: null,
        location_source: "microsoft",
        location_lookup: "confirmed",
      }),
    ]);
    const [finding] = await findings();
    expect(String(fieldsOf(finding)?.signals)).toBe(
      "new.ip, new.country, new.state, travel.impossible",
    );
    expect(limitationsOf(finding).join(" ")).toContain("agrees with Microsoft's location");
  });

  it("keeps Microsoft's names when the provider places the sign-in nearby", async () => {
    // The same city under the provider's English name, 2 km from Microsoft's point.
    const provider = fakeProvider(() => ({
      ...chicagoAnswer,
      country: "NL",
      state: "North Holland",
      city: "Amsterdam",
      lat: 52.38,
      lon: 4.88,
    }));
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77"));
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({
        state: "Noord-Holland",
        location_source: "microsoft",
        location_lookup: "confirmed",
      }),
    ]);
  });

  it("never lets a country-level answer overrule Microsoft's place in that country", async () => {
    // Microsoft says Seattle; the provider knows only "United States", centred about 2,300 km
    // away. Nothing contradicts Seattle, so Microsoft's location stands.
    const provider = fakeProvider(() => ({
      country: "US",
      state: null,
      city: null,
      lat: 37.75,
      lon: -97.82,
      accuracyKm: 1_000,
      network: null,
    }));
    await setup(provider);
    await poll();
    records.push(
      ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77", {
        location: {
          city: "Seattle",
          state: "Washington",
          countryOrRegion: "US",
          geoCoordinates: { latitude: 47.61, longitude: -122.33 },
        },
      }),
    );
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({
        country: "US",
        state: "Washington",
        city: "Seattle",
        lat: 47.61,
        lon: -122.33,
        location_source: "microsoft",
        location_lookup: "confirmed",
      }),
    ]);
  });

  it("flags a real move to a new region of the account's own country, found by the provider", async () => {
    // Microsoft says Seattle; the provider says Milwaukee, about 130 km from the account's
    // Chicago: too close for travel, but a region this account has never signed in from.
    const provider = fakeProvider(() => ({
      ...chicagoAnswer,
      state: "Wisconsin",
      city: "Milwaukee",
      lat: 43.04,
      lon: -87.91,
    }));
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77", { location: seattle }));
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({ state: "Wisconsin", location_lookup: "found" }),
    ]);
    const [finding] = await findings();
    expect(String(fieldsOf(finding)?.signals)).toBe("new.ip, new.state");
  });

  it("knows a region the provider names differently when the account has been there", async () => {
    // Chicago again, under a name Microsoft never uses. Microsoft's own sign-ins from Chicago
    // are within reach, so the region is familiar and nothing is raised.
    const provider = fakeProvider(() => ({ ...chicagoAnswer, state: "State of Illinois" }));
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77", { location: seattle }));
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({ state: "State of Illinois", location_lookup: "found" }),
    ]);
    expect(await findings()).toEqual([]);
  });

  it("takes the provider's precise place over Microsoft's a city away", async () => {
    // Microsoft says Milwaukee, about 130 km from the provider's Chicago (radius 20 km): beyond
    // the 25 km a single city spans, so the provider's place is used.
    const provider = fakeProvider(() => chicagoAnswer);
    await setup(provider);
    await poll();
    records.push(
      ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77", {
        location: {
          city: "Milwaukee",
          state: "Wisconsin",
          countryOrRegion: "US",
          geoCoordinates: { latitude: 43.04, longitude: -87.91 },
        },
      }),
    );
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({
        city: "Chicago",
        ms_city: "Milwaukee",
        location_lookup: "found",
      }),
    ]);
    expect(await findings()).toEqual([]);
  });

  it("keeps Microsoft's location when it gave no point and the provider agrees on the country", async () => {
    const provider = fakeProvider(() => ({
      ...chicagoAnswer,
      country: "NL",
      state: "Limburg",
      city: "Maastricht",
      lat: 50.85,
      lon: 5.69,
    }));
    await setup(provider);
    await poll();
    records.push(
      ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77", {
        location: { city: "Maastricht", state: "Limburg", countryOrRegion: "NL" },
      }),
    );
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({
        state: "Limburg",
        ms_state: null,
        location_source: "microsoft",
        location_lookup: "confirmed",
      }),
    ]);
    const [finding] = await findings();
    expect(String(fieldsOf(finding)?.signals)).toBe("new.ip, new.country, new.state");
  });

  it("reuses an answer for every network inside the provider's own range", async () => {
    const provider = fakeProvider(() => ({ ...chicagoAnswer, network: "2001:db8:1::/48" }));
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77"));
    await poll();
    // Another device, another /64, same /48: answered from the cache.
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:99::5"));
    await poll();
    expect(provider.calls).toEqual([["2001:db8:1:2::/64"]]);
    expect(await stored("2001:db8:1:99::5")).toEqual([
      expect.objectContaining({
        country: "US",
        location_source: "test-geo",
        location_lookup: "found",
      }),
    ]);
    const cache = await database.tenant(tenantId, (sql) =>
      sql.query<{ network: string }>("SELECT network FROM mod_signins_audit_ip_location"),
    );
    // Networks only: no full address is ever cached.
    expect(cache.rows).toEqual([{ network: "2001:db8:1::/48" }]);
  });

  it("never asks about a network again once it has an answer, however old, and never deletes it", async () => {
    const provider = fakeProvider((network) =>
      network === "2001:db8:1:2::/64" ? { ...chicagoAnswer, network } : null,
    );
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77"));
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:dead:beef::1"));
    await poll();
    expect(provider.calls.flat().sort()).toEqual(["2001:db8:1:2::/64", "2001:db8:dead:beef::/64"]);
    // Both answers, found and "no data", saved over a year ago.
    await database.tenant(tenantId, (sql) =>
      sql.query(
        "UPDATE mod_signins_audit_ip_location SET looked_up_at = looked_up_at - interval '400 days'",
      ),
    );
    records.push(ipv6SignIn(minutesAgo(1), "2001:db8:1:2::78"));
    records.push(ipv6SignIn(minutesAgo(1), "2001:db8:dead:beef::2"));
    await poll();
    expect(provider.calls).toHaveLength(1);
    expect(await stored("2001:db8:1:2::78")).toEqual([
      expect.objectContaining({ location_source: "test-geo", location_lookup: "found" }),
    ]);
    expect(await stored("2001:db8:dead:beef::2")).toEqual([
      expect.objectContaining({ location_source: "microsoft", location_lookup: "no_data" }),
    ]);
    // No retention rule reaches the saved answers.
    const retention = createSigninsAuditModule().retention.map((rule) => rule.table);
    expect(retention).not.toContain("ip_location");
  });

  it("keeps Microsoft's location for a network the provider cannot place, so detection still works", async () => {
    const provider = fakeProvider(() => null);
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:dead:beef::1"));
    await poll();
    expect(await stored("2001:db8:dead:beef::1")).toEqual([
      expect.objectContaining({
        country: "NL",
        location_source: "microsoft",
        location_lookup: "no_data",
      }),
    ]);
    const [finding] = await findings();
    expect(limitationsOf(finding).join(" ")).toContain("has no location for");
  });

  it("falls back to Microsoft's location when the provider cannot answer, without failing the run", async () => {
    const provider = fakeProvider(() => chicagoAnswer, {
      fail: new IpLocationUnavailableError("throttled"),
    });
    await setup(provider);
    await poll();
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77"));
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({
        country: "NL",
        location_source: "microsoft",
        location_lookup: "unavailable",
        processed: true,
      }),
    ]);
    const [run] = await database.tenant(tenantId, (sql) => latestRuns(sql, tenantId));
    expect(run?.status).toBe("succeeded");
    expect(limitationsOf((await findings())[0]).join(" ")).toContain("could not be asked in time");
  });

  it("stops asking once the run's lookup time is spent and keeps Microsoft's location for the rest", async () => {
    // Each request takes 16 seconds of the run's 15-second lookup budget.
    const provider = fakeProvider(() => chicagoAnswer, { delayMs: 16_000 });
    await setup(provider);
    await poll();
    for (let index = 1; index <= 25; index += 1)
      records.push(ipv6SignIn(minutesAgo(3), `2001:db8:${index.toString(16)}::1`));
    await poll();
    expect(provider.calls).toHaveLength(1);
    expect(provider.calls[0]).toHaveLength(8);
    const outcomes = await database.tenant(tenantId, (sql) =>
      sql.query<{ location_lookup: string; n: number }>(
        `SELECT location_lookup, count(*)::int AS n FROM mod_signins_audit_event
          WHERE ip LIKE '2001:db8:%' GROUP BY 1 ORDER BY 1`,
      ),
    );
    expect(outcomes.rows).toEqual([
      { location_lookup: "found", n: 8 },
      { location_lookup: "skipped", n: 17 },
    ]);
    expect(
      await database.tenant(tenantId, (sql) =>
        sql.query("SELECT 1 FROM mod_signins_audit_event WHERE processed = false"),
      ),
    ).toMatchObject({ rows: [] });
  });

  it("drops coarse coordinates, so a country's centre never reads as impossible travel", async () => {
    // Microsoft says Amsterdam; the provider says only "United States", at its centre.
    const provider = fakeProvider(() => ({
      country: "US",
      state: null,
      city: null,
      lat: 37.75,
      lon: -97.82,
      accuracyKm: 1_000,
      network: null,
    }));
    await setup(provider);
    await poll();
    // Chicago eight minutes earlier, then a new device: about 750 km to the centre in eight
    // minutes would be impossible travel if those coordinates were used.
    records.push(signIn(minutesAgo(10)));
    records.push(ipv6SignIn(minutesAgo(2), "2001:db8:1:2::77"));
    await poll();
    expect(await stored("2001:db8:1:2::77")).toEqual([
      expect.objectContaining({ country: "US", lat: null, lon: null, location_lookup: "found" }),
    ]);
    expect(await findings()).toEqual([]);
  });

  it("never sends failed sign-ins or addresses that carry IPv4", async () => {
    const provider = fakeProvider(() => chicagoAnswer);
    await setup(provider);
    await poll();
    records.push(
      ipv6SignIn(minutesAgo(3), "2001:db8:1:2::77", {
        status: { errorCode: 50126, failureReason: "Invalid username or password." },
      }),
    );
    // 6to4: the network part is the IPv4 address 192.0.2.10.
    records.push(ipv6SignIn(minutesAgo(2), "2002:c000:20a::1"));
    await poll();
    expect(provider.calls).toEqual([]);
    for (const ip of ["2001:db8:1:2::77", "2002:c000:20a::1"])
      expect(await stored(ip)).toEqual([
        expect.objectContaining({ location_source: "microsoft", location_lookup: null }),
      ]);
  });

  it("never looks up a failed sign-in left unprocessed by a run before this change", async () => {
    // A build without lookups stored this failed attempt and stopped before detection, as a run
    // cut short at upgrade time does: not processed, and no location source, like every row
    // that build writes. The upgraded run must neither send its network nor move it.
    const provider = fakeProvider(() => chicagoAnswer);
    await setup(provider);
    await poll();
    const ip = "2001:db8:1:2::77";
    records.push(
      ipv6SignIn(minutesAgo(3), ip, {
        status: { errorCode: 50126, failureReason: "Invalid username or password." },
      }),
    );
    await poll();
    await database.tenant(tenantId, (sql) =>
      sql.query(
        "UPDATE mod_signins_audit_event SET processed = false, location_source = NULL WHERE ip = $1",
        [ip],
      ),
    );
    await poll();
    expect(provider.calls).toEqual([]);
    expect(await stored(ip)).toEqual([
      expect.objectContaining({
        country: amsterdam.countryOrRegion,
        location_source: null,
        location_lookup: null,
        processed: true,
      }),
    ]);
  });

  it("still sees a dormant account whose last sign-in had no coordinates", async () => {
    // No provider at all: this is about the rules every installation runs.
    records = [
      signIn(new Date(start.getTime() - 20 * 86_400_000), {
        location: { city: "Chicago", state: "Illinois", countryOrRegion: "US" },
      }),
    ];
    await setup(undefined, { backfillDays: 25, dormantDays: 14 });
    await poll();
    records.push(
      signIn(minutesAgo(2), {
        ipAddress: "192.0.2.77",
        deviceDetail: { deviceId: "", operatingSystem: "Linux", browser: "Firefox" },
      }),
    );
    await poll();
    const [finding] = await findings();
    expect(String(fieldsOf(finding)?.signals)).toBe("new.ip, dormant.revival");
  });

  it("measures travel from the last sign-in that has coordinates", async () => {
    const provider = fakeProvider(() => ({
      country: "US",
      state: null,
      city: null,
      lat: null,
      lon: null,
      accuracyKm: null,
      network: null,
    }));
    await setup(provider);
    await poll();
    // Chicago (IPv4) two hours ago from history, then, read in the same window, an IPv6 sign-in
    // with no coordinates and three minutes later Amsterdam.
    records.push(
      ipv6SignIn(minutesAgo(4), "2001:db8:1:2::77", {
        deviceDetail: { deviceId: "device-1", operatingSystem: "Windows", browser: "Edge" },
      }),
    );
    records.push(
      signIn(minutesAgo(1), {
        ipAddress: "192.0.2.99",
        location: amsterdam,
        deviceDetail: { deviceId: "", operatingSystem: "Linux", browser: "Chrome" },
      }),
    );
    await poll();
    const [finding] = await findings();
    expect(String(fieldsOf(finding)?.signals)).toContain("travel.impossible");
  });
});
