import type { PGliteModuleDatabase } from "@prism365/storage";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  type CollectionCredential,
  containingNetworks,
  type IpLocation,
  type IpLocationProvider,
  type IpLocationRequest,
  IpLocationUnavailableError,
  ipv6Network64,
  ipv6Prefix,
  isIpv6Network64,
  latestRuns,
  ModuleDefinitionError,
  ModuleRegistry,
  ModuleScheduler,
  type PrismModule,
  providerNetwork,
} from "../src/index.ts";
import {
  actor,
  fakeGraph,
  openRuntime,
  type ProbeBehavior,
  probeModule,
  tenantA,
} from "./fixtures.ts";

// Addresses are documentation ranges (2001:db8::/32, 192.0.2.0/24).

// The boundary tests open PGlite each time; on a loaded machine that cold start alone can pass
// the default 5 s budget.
vi.setConfig({ testTimeout: 30_000 });

describe("IPv6 networks for lookups", () => {
  it("keeps only the /64 of a global unicast address, in one written form", () => {
    expect(ipv6Network64("2001:db8:1:2:3:4:5:6")).toBe("2001:db8:1:2::/64");
    expect(ipv6Network64("2001:0DB8:0000:0001::1")).toBe("2001:db8:0:1::/64");
    expect(ipv6Network64(" 2001:db8:: ")).toBe("2001:db8:0:0::/64");
    expect(ipv6Network64("2001:db8:a:b:c:d:e:f")).toBe(ipv6Network64("2001:db8:a:b::1"));
  });

  it("refuses everything that is not global unicast IPv6", () => {
    for (const value of [
      null,
      "",
      "192.0.2.10",
      "::ffff:192.0.2.10",
      "::1",
      "::",
      "fe80::1",
      "fe80::1%eth0",
      "fc00::1",
      "fd00:ec2::254",
      "ff02::1",
      "2001:db8::g",
      "2001:db8:::1",
      "not an address",
      // 6to4 and Teredo carry an IPv4 address in the network part.
      "2002:c000:20a::1",
      "2001:0:4136:e378:8000:63bf:3fff:fdd2",
    ])
      expect(ipv6Network64(value)).toBeNull();
  });

  it("recognises only networks it could have written", () => {
    expect(isIpv6Network64("2001:db8:1:2::/64")).toBe(true);
    for (const value of [
      "2001:db8:1:2::1/64",
      "2001:db8:01:2::/64",
      "2001:DB8:1:2::/64",
      "2001:db8::/64",
      "2001:db8:1:2::/48",
      "2001:db8:1:2::",
      "fe80:0:0:0::/64",
      "2002:c000:20a:0::/64",
      "192.0.2.0/24",
      42,
    ])
      expect(isIpv6Network64(value)).toBe(false);
  });

  it("masks a network at any prefix length up to /64", () => {
    expect(ipv6Prefix("2001:db8:1:2abc::", 56)).toBe("2001:db8:1:2a00::/56");
    expect(ipv6Prefix("2001:db8:1:2abc::", 48)).toBe("2001:db8:1::/48");
    expect(ipv6Prefix("2001:db8:ffff::", 29)).toBe("2001:db8::/29");
    expect(ipv6Prefix("2001:db8::", 65)).toBeNull();
    expect(ipv6Prefix("fe80::", 16)).toBeNull();
  });

  it("files a provider's range only when it holds the asked network, clamped to /32../64", () => {
    const network = "2001:db8:1:2::/64";
    expect(providerNetwork("2001:db8::/32", network)).toBe("2001:db8::/32");
    expect(providerNetwork("2001:db8:1::/48", network)).toBe("2001:db8:1::/48");
    expect(providerNetwork("2000::/12", network)).toBe("2001:db8::/32");
    expect(providerNetwork("2001:db8:1:2:3::/80", network)).toBe(network);
    expect(providerNetwork("2001:db9::/32", network)).toBeNull();
    expect(providerNetwork("192.0.2.0/24", network)).toBeNull();
    expect(providerNetwork("2001:db8::/129", network)).toBeNull();
    expect(providerNetwork(undefined, network)).toBeNull();
  });

  it("lists every /64 to /32 network that could hold a cached answer, longest first", () => {
    const networks = containingNetworks("2001:db8:1:2::/64");
    expect(networks).toHaveLength(33);
    expect(networks[0]).toBe("2001:db8:1:2::/64");
    expect(networks.at(-1)).toBe("2001:db8::/32");
    expect(networks).toContain("2001:db8:1::/48");
    expect(containingNetworks("2001:db8:1:2::1/64")).toEqual([]);
  });
});

describe("IP location at the scheduler boundary", () => {
  let database: PGliteModuleDatabase | undefined;
  let scheduler: ModuleScheduler | undefined;

  afterEach(async () => {
    await scheduler?.stop();
    await database?.close();
    scheduler = undefined;
    database = undefined;
  });

  function provider(answer: (networks: readonly string[]) => Promise<void> = async () => {}) {
    const calls: { networks: string[]; request: IpLocationRequest }[] = [];
    const location: IpLocation = {
      country: "XA",
      state: null,
      city: null,
      lat: null,
      lon: null,
      accuracyKm: null,
      network: null,
    };
    const value: IpLocationProvider = {
      id: "test-geo",
      name: "Synthetic lookup",
      attribution: null,
      locate: async (networks, request) => {
        calls.push({ networks: [...networks], request });
        await answer(networks);
        // An extra key the module did not ask about is dropped by the scheduler.
        return new Map<string, IpLocation | null>([
          ...networks.map((network) => [network, location] as const),
          ["2001:db8:ffff:ffff::/64", location],
        ]);
      },
    };
    return { calls, value };
  }

  async function run(
    behavior: ProbeBehavior,
    options: { enrichment?: PrismModule["enrichment"]; ipLocation?: IpLocationProvider } = {},
    credential: Partial<CollectionCredential> = {},
  ) {
    const module: PrismModule = {
      ...probeModule({ run: behavior }),
      ...(options.enrichment ? { enrichment: options.enrichment } : {}),
    };
    let registry: ModuleRegistry;
    ({ registry, database } = await openRuntime(module));
    scheduler = new ModuleScheduler({
      enabledPacks: ["directory.users.read"],
      registry,
      database,
      graphClient: fakeGraph().factory,
      now: () => new Date("2026-11-09T12:00:00Z"),
      ...(options.ipLocation ? { ipLocation: options.ipLocation } : {}),
    });
    scheduler.bind({
      kind: "delegated",
      tenantId: tenantA,
      actorId: actor,
      grantedPermissions: async () => ["User.Read", "User.Read.All"],
      accessToken: async () => "synthetic-token",
      ...credential,
    });
    await scheduler.tick();
    const db = database;
    return (await db.tenant(tenantA, (sql) => latestRuns(sql, tenantA)))[0];
  }

  it("is offered only to a module that declared it, and only when a provider is configured", async () => {
    const seen: boolean[] = [];
    const record: ProbeBehavior = async (context) => {
      seen.push(context.enrichment.ipLocation !== null);
      return undefined;
    };
    const { value } = provider();
    const reset = async () => {
      await scheduler?.stop();
      await database?.close();
      scheduler = undefined;
      database = undefined;
    };
    await run(record, { ipLocation: value });
    await reset();
    await run(record, { enrichment: ["ip.location"] });
    await reset();
    await run(record, { enrichment: ["ip.location"], ipLocation: value });
    expect(seen).toEqual([false, false, true]);
  });

  it("sends only checked /64 networks and returns only what was asked", async () => {
    const { calls, value } = provider();
    const answers: unknown[] = [];
    const refused: string[] = [];
    const result = await run(
      async (context) => {
        const lookup = context.enrichment.ipLocation;
        if (!lookup) throw new Error("no lookup");
        expect(lookup.provider).toEqual({
          id: "test-geo",
          name: "Synthetic lookup",
          attribution: null,
        });
        answers.push([...(await lookup.locate(["2001:db8:1:2::/64"])).entries()]);
        for (const bad of [
          ["2001:db8:1:2::1"],
          ["2001:db8:1:2::5/64"],
          ["192.0.2.0/24"],
          ["fe80:0:0:0::/64"],
          [],
          ["2001:db8:1:2::/64", "2001:db8:1:2::/64"],
          Array.from({ length: 101 }, (_value, index) => `2001:db8:${index.toString(16)}:0::/64`),
        ])
          await lookup.locate(bad).catch((error: Error) => refused.push(error.name));
        return undefined;
      },
      { enrichment: ["ip.location"], ipLocation: value },
    );
    expect(result).toMatchObject({ status: "succeeded" });
    expect(refused).toEqual(Array.from({ length: 7 }, () => "ModuleBoundaryError"));
    expect(calls.map((call) => call.networks)).toEqual([["2001:db8:1:2::/64"]]);
    expect(calls[0]?.request).toMatchObject({
      tenantId: tenantA,
      actorId: actor,
      moduleId: "test.probe",
    });
    expect(answers).toEqual([[["2001:db8:1:2::/64", expect.objectContaining({ country: "XA" })]]]);
  });

  it("lets a provider failure reach the module as IpLocationUnavailableError", async () => {
    const { value } = provider(async () => {
      throw new IpLocationUnavailableError("quota");
    });
    const reasons: string[] = [];
    await run(
      async (context) => {
        await context.enrichment.ipLocation
          ?.locate(["2001:db8:1:2::/64"])
          .catch((error: IpLocationUnavailableError) => reasons.push(error.reason));
        return undefined;
      },
      { enrichment: ["ip.location"], ipLocation: value },
    );
    expect(reasons).toEqual(["quota"]);
  });

  it("refuses a module that declares an unknown or repeated enrichment", () => {
    const base = probeModule({ run: async () => undefined });
    expect(() =>
      new ModuleRegistry().register({
        ...base,
        enrichment: ["ip.reputation" as "ip.location"],
      }),
    ).toThrow(ModuleDefinitionError);
    expect(() =>
      new ModuleRegistry().register({ ...base, enrichment: ["ip.location", "ip.location"] }),
    ).toThrow(ModuleDefinitionError);
  });
});
