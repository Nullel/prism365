import { randomUUID } from "node:crypto";
import { IpLocationUnavailableError } from "@prism365/modules";
import { describe, expect, it } from "vitest";
import {
  type IpLocationConfig,
  ipLocationConfigFromEnv,
  MaxMindIpLocation,
} from "../src/ip-location.ts";

// The IP location provider for IPv6 sign-ins. Every test uses a fake fetch; none opens a
// connection. Networks are documentation ranges; the account and key are invented.

const files = (content: Record<string, string>) => async (path: string) => {
  const value = content[path];
  if (value === undefined) throw Object.assign(new Error("missing"), { code: "ENOENT" });
  return value;
};

const key = "synthetic_license_key_0000000000000000_mmk";

describe("IP location configuration", () => {
  const read = files({ "/run/secrets/maxmind": `${key}\n` });

  it("is off unless a provider is named", async () => {
    expect(await ipLocationConfigFromEnv({}, read)).toBeUndefined();
  });

  it("reads the license key from a file only", async () => {
    expect(
      await ipLocationConfigFromEnv(
        {
          IP_LOCATION_PROVIDER: "maxmind-geolite",
          IP_LOCATION_ACCOUNT_ID: "123456",
          IP_LOCATION_LICENSE_KEY_FILE: "/run/secrets/maxmind",
        },
        read,
      ),
    ).toEqual({ provider: "maxmind-geolite", accountId: "123456", licenseKey: key });
    await expect(ipLocationConfigFromEnv({ IP_LOCATION_LICENSE_KEY: key }, read)).rejects.toThrow(
      /never read/,
    );
  });

  it("refuses a half or wrong configuration, without echoing the key", async () => {
    const base = {
      IP_LOCATION_PROVIDER: "maxmind-geolite",
      IP_LOCATION_ACCOUNT_ID: "123456",
      IP_LOCATION_LICENSE_KEY_FILE: "/run/secrets/maxmind",
    };
    for (const env of [
      { IP_LOCATION_ACCOUNT_ID: "123456" },
      { IP_LOCATION_LICENSE_KEY_FILE: "/run/secrets/maxmind" },
      { ...base, IP_LOCATION_PROVIDER: "ipinfo" },
      { ...base, IP_LOCATION_ACCOUNT_ID: "abc" },
      { ...base, IP_LOCATION_ACCOUNT_ID: "0123" },
      { ...base, IP_LOCATION_LICENSE_KEY_FILE: "" },
      { ...base, IP_LOCATION_LICENSE_KEY_FILE: "/missing" },
    ]) {
      const error = await ipLocationConfigFromEnv(env, read).catch((caught: Error) => caught);
      expect(error).toBeInstanceOf(Error);
      expect(String((error as Error).message)).not.toContain(key);
    }
    await expect(
      ipLocationConfigFromEnv(base, files({ "/run/secrets/maxmind": "has spaces in it\n" })),
    ).rejects.toThrow(/does not hold a license key/);
  });
});

interface Sent {
  url: string;
  init: RequestInit | undefined;
}

const config: IpLocationConfig = {
  provider: "maxmind-geolite",
  accountId: "123456",
  licenseKey: key,
};

const json = (status: number, body: unknown, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...headers },
  });

const city = {
  city: { names: { en: "Chicago" } },
  country: { iso_code: "us", names: { en: "United States" } },
  location: { accuracy_radius: 20, latitude: 41.88, longitude: -87.63 },
  subdivisions: [{ iso_code: "IL", names: { en: "Illinois" } }],
  traits: { ip_address: "2001:db8:1:2::", network: "2001:db8:1::/48" },
};

function harness(answer: (url: string) => Response | Promise<Response>) {
  const sent: Sent[] = [];
  const ledger: { tenantId: string; actorId: string; event: Record<string, unknown> }[] = [];
  const order: string[] = [];
  const paused: string[] = [];
  let now = Date.parse("2026-11-09T12:00:00Z");
  const provider = new MaxMindIpLocation(config, {
    ledger: async (tenantId, actorId, event) => {
      order.push("ledger");
      ledger.push({ tenantId, actorId, event });
    },
    fetchImplementation: async (input, init) => {
      order.push("fetch");
      sent.push({ url: String(input), init });
      return answer(String(input));
    },
    now: () => now,
    onPause: (reason) => paused.push(reason),
  });
  const request = {
    tenantId: randomUUID(),
    actorId: randomUUID(),
    moduleId: "signins.audit",
    signal: new AbortController().signal,
  };
  return {
    provider,
    request,
    sent,
    ledger,
    order,
    paused,
    advance: (milliseconds: number) => {
      now += milliseconds;
    },
  };
}

const reason = (promise: Promise<unknown>) =>
  promise.then(
    () => "answered",
    (error: unknown) =>
      error instanceof IpLocationUnavailableError ? error.reason : `other: ${String(error)}`,
  );

describe("MaxMind IP location", () => {
  it("records the lookup in the access ledger before sending, with counts only", async () => {
    const h = harness(() => json(200, city));
    await h.provider.locate(["2001:db8:1:2::/64", "2001:db8:0:0::/64"], h.request);
    expect(h.order).toEqual(["ledger", "fetch", "fetch"]);
    expect(h.ledger).toEqual([
      {
        tenantId: h.request.tenantId,
        actorId: h.request.actorId,
        event: {
          type: "enrichment.ip_location",
          provider: "maxmind-geolite",
          module: "signins.audit",
          networks: 2,
          occurredAt: "2026-11-09T12:00:00.000Z",
        },
      },
    ]);
    expect(JSON.stringify(h.ledger)).not.toContain("2001:db8");
  });

  it("asks geolite.info for each network's first address, with HTTP Basic and no redirects", async () => {
    const h = harness(() => json(200, city));
    await h.provider.locate(["2001:db8:1:2::/64", "2001:db8:0:0::/64"], h.request);
    expect(h.sent.map((request) => request.url).sort()).toEqual([
      "https://geolite.info/geoip/v2.1/city/2001:db8:1:2::",
      "https://geolite.info/geoip/v2.1/city/2001:db8::",
    ]);
    const headers = new Headers(h.sent[0]?.init?.headers);
    expect(headers.get("authorization")).toBe(
      `Basic ${Buffer.from(`123456:${key}`).toString("base64")}`,
    );
    expect(h.sent[0]?.init?.redirect).toBe("manual");
    expect(h.sent[0]?.init?.method).toBe("GET");
  });

  it("uses geoip.maxmind.com for the paid service", async () => {
    const sent: string[] = [];
    const provider = new MaxMindIpLocation(
      { ...config, provider: "maxmind-geoip" },
      {
        ledger: async () => {},
        fetchImplementation: async (input) => {
          sent.push(String(input));
          return json(200, city);
        },
      },
    );
    expect(provider.attribution).toBeNull();
    await provider.locate(["2001:db8:1:2::/64"], {
      tenantId: randomUUID(),
      actorId: randomUUID(),
      moduleId: "signins.audit",
      signal: new AbortController().signal,
    });
    expect(sent).toEqual(["https://geoip.maxmind.com/geoip/v2.1/city/2001:db8:1:2::"]);
  });

  it("reads a city answer, including the record's own network", async () => {
    const h = harness(() => json(200, city));
    const answers = await h.provider.locate(["2001:db8:1:2::/64"], h.request);
    expect(answers.get("2001:db8:1:2::/64")).toEqual({
      country: "US",
      state: "Illinois",
      city: "Chicago",
      lat: 41.88,
      lon: -87.63,
      accuracyKm: 20,
      network: "2001:db8:1::/48",
    });
    expect(h.provider.attribution).toContain("GeoLite Data created by MaxMind");
  });

  it("drops a record network that does not hold the asked network", async () => {
    const h = harness(() => json(200, { ...city, traits: { network: "2001:db9::/32" } }));
    const answers = await h.provider.locate(["2001:db8:1:2::/64"], h.request);
    expect(answers.get("2001:db8:1:2::/64")?.network).toBeNull();
  });

  it("answers no data for an address MaxMind does not have or calls reserved", async () => {
    const h = harness((url) =>
      url.endsWith("2001:db8:1:2::")
        ? json(404, { code: "IP_ADDRESS_NOT_FOUND", error: "not in the database" })
        : json(400, { code: "IP_ADDRESS_RESERVED", error: "reserved" }),
    );
    const answers = await h.provider.locate(["2001:db8:1:2::/64", "2001:db8:0:0::/64"], h.request);
    expect([...answers.values()]).toEqual([null, null]);
    expect(h.paused).toEqual([]);
  });

  it("pauses for an hour when the daily quota is spent, without asking or recording again", async () => {
    const h = harness(() => json(402, { code: "INSUFFICIENT_FUNDS", error: "out" }));
    expect(await reason(h.provider.locate(["2001:db8:1:2::/64"], h.request))).toBe("quota");
    expect(h.paused).toEqual(["quota"]);
    h.advance(59 * 60_000);
    expect(await reason(h.provider.locate(["2001:db8:1:2::/64"], h.request))).toBe("quota");
    expect(h.sent).toHaveLength(1);
    expect(h.ledger).toHaveLength(1);
    h.advance(2 * 60_000);
    await reason(h.provider.locate(["2001:db8:1:2::/64"], h.request));
    expect(h.sent).toHaveLength(2);
  });

  it("maps each failure to a fixed reason", async () => {
    const cases: [() => Response | Promise<Response>, string][] = [
      [() => json(401, { code: "AUTHORIZATION_INVALID", error: "x" }), "refused"],
      [() => json(403, { code: "PERMISSION_REQUIRED", error: "x" }), "refused"],
      [() => json(429, {}, { "Retry-After": "120" }), "throttled"],
      [() => json(503, {}), "unavailable"],
      [() => json(404, { code: "SOMETHING_ELSE", error: "x" }), "unavailable"],
      [
        () =>
          new Response(null, { status: 301, headers: { Location: "https://example.invalid/" } }),
        "refused",
      ],
      // A 200 in a shape MaxMind does not document is never taken as "no data".
      [() => json(200, { country: { iso_code: 42 } }), "refused"],
      [() => new Response("not json", { status: 200 }), "refused"],
      [() => new Response("x".repeat(70 * 1024), { status: 200 }), "refused"],
      [
        () => {
          throw new TypeError("fetch failed");
        },
        "unavailable",
      ],
    ];
    for (const [answer, expected] of cases) {
      const h = harness(answer);
      expect(await reason(h.provider.locate(["2001:db8:1:2::/64"], h.request))).toBe(expected);
    }
  });

  it("waits out Retry-After before asking again", async () => {
    const h = harness(() => json(429, {}, { "Retry-After": "120" }));
    await reason(h.provider.locate(["2001:db8:1:2::/64"], h.request));
    h.advance(119_000);
    expect(await reason(h.provider.locate(["2001:db8:1:2::/64"], h.request))).toBe("throttled");
    expect(h.sent).toHaveLength(1);
    h.advance(2_000);
    await reason(h.provider.locate(["2001:db8:1:2::/64"], h.request));
    expect(h.sent).toHaveLength(2);
  });

  it("sends nothing it cannot record, and nothing that is not a /64 network", async () => {
    const noActor = harness(() => json(200, city));
    expect(
      await reason(
        noActor.provider.locate(["2001:db8:1:2::/64"], { ...noActor.request, actorId: null }),
      ),
    ).toBe("ledger");
    expect(noActor.sent).toEqual([]);

    const failingLedger = new MaxMindIpLocation(config, {
      ledger: async () => {
        throw new Error("ledger down");
      },
      fetchImplementation: async () => {
        throw new Error("must not be called");
      },
    });
    expect(
      await reason(
        failingLedger.locate(["2001:db8:1:2::/64"], {
          tenantId: randomUUID(),
          actorId: randomUUID(),
          moduleId: "signins.audit",
          signal: new AbortController().signal,
        }),
      ),
    ).toBe("ledger");

    const h = harness(() => json(200, city));
    for (const bad of [["2001:db8:1:2::1"], ["192.0.2.0/24"], []])
      expect(await reason(h.provider.locate(bad, h.request))).toBe("refused");
    expect(h.sent).toEqual([]);
    expect(h.ledger).toEqual([]);
  });

  it("sends nothing when the run ends while the ledger entry is written", async () => {
    const run = new AbortController();
    const sent: string[] = [];
    const provider = new MaxMindIpLocation(config, {
      ledger: async () => {
        run.abort();
      },
      fetchImplementation: async (input) => {
        sent.push(String(input));
        return json(200, city);
      },
    });
    expect(
      await reason(
        provider.locate(["2001:db8:1:2::/64", "2001:db8:1:3::/64"], {
          tenantId: randomUUID(),
          actorId: randomUUID(),
          moduleId: "signins.audit",
          signal: run.signal,
        }),
      ),
    ).toBe("timeout");
    expect(sent).toEqual([]);
  });

  it("stops the requests still waiting once one fails", async () => {
    const h = harness(() => json(503, {}));
    const networks = Array.from({ length: 12 }, (_value, index) => `2001:db8:${index + 1}:0::/64`);
    expect(await reason(h.provider.locate(networks, h.request))).toBe("unavailable");
    expect(h.sent.length).toBeLessThanOrEqual(4);
  });
});
