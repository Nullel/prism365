import { describe, expect, it } from "vitest";
import {
  classifyAddress,
  type HostLookup,
  parseWebhookUrl,
  resolvePublicAddress,
} from "../src/delivery-address.ts";

// Where a webhook may point. The public addresses below sit just outside the blocked ranges, so
// each range's edge is tested from both sides; no test connects to anything.

describe("webhook addresses as typed", () => {
  const parse = (raw: string, ports: number[] = []) => parseWebhookUrl(raw, ports);

  it("takes an HTTPS address with a host name and keeps its path and query", () => {
    const result = parse("https://hooks.example/services/T000/B000/synthetic?sig=abc#ignored");
    expect(result).toEqual({
      ok: true,
      url: {
        href: "https://hooks.example/services/T000/B000/synthetic?sig=abc",
        host: "hooks.example",
        port: 443,
        path: "/services/T000/B000/synthetic?sig=abc",
      },
    });
    // The default port written out is still the default port.
    expect(parse("https://hooks.example:443/x")).toMatchObject({ ok: true, url: { port: 443 } });
  });

  it("refuses anything but HTTPS", () => {
    for (const raw of [
      "http://hooks.example/x",
      "ftp://hooks.example/x",
      "file:///etc/passwd",
      "gopher://hooks.example/x",
      "//hooks.example/x",
      "hooks.example/x",
    ])
      expect(parse(raw), raw).toEqual({ ok: false, problem: "invalid_url" });
  });

  it("refuses user info in every spelling", () => {
    for (const raw of [
      "https://user@hooks.example/x",
      "https://user:secret@hooks.example/x",
      "https://:secret@hooks.example/x",
      "https://hooks.example@internal.example/x",
      "https://hooks.example%40internal.example/x",
    ])
      expect(parse(raw), raw).toEqual({ ok: false, problem: "invalid_url" });
  });

  it("refuses address literals, whatever they point to", () => {
    for (const raw of [
      "https://127.0.0.1/x",
      "https://100.128.0.1/x",
      "https://[::1]/x",
      "https://[2003::1]/x",
      "https://0x7f.1/x",
      "https://2130706433/x",
      "https://017700000001/x",
      "https://127.1/x",
    ])
      expect(parse(raw), raw).toEqual({ ok: false, problem: "invalid_url" });
  });

  it("refuses local names", () => {
    for (const raw of [
      "https://localhost/x",
      "https://api.localhost/x",
      "https://printer.local/x",
      "https://metadata.google.internal/x",
      "https://1.0.0.127.in-addr.arpa/x",
      "https://intranet/x",
      "https://hooks.example./x",
    ])
      expect(parse(raw), raw).toEqual({ ok: false, problem: "invalid_url" });
  });

  it("refuses whitespace, control characters and backslashes the parser would drop", () => {
    for (const raw of [
      "https://hooks.exa\nmple/x",
      "https://hooks.exa\tmple/x",
      " https://hooks.example/x",
      "https://hooks.example\\@internal.example/x",
      "https://hooks.example/x\u200b",
      `https://hooks.example/${"a".repeat(2_100)}`,
    ])
      expect(parse(raw), JSON.stringify(raw)).toEqual({ ok: false, problem: "invalid_url" });
  });

  it("allows another port only when WEBHOOK_ALLOWED_PORTS lists it", () => {
    expect(parse("https://hooks.example:8443/x")).toEqual({
      ok: false,
      problem: "port_not_allowed",
    });
    expect(parse("https://hooks.example:8443/x", [8443])).toMatchObject({
      ok: true,
      url: { port: 8443 },
    });
    expect(parse("https://hooks.example:22/x", [8443])).toEqual({
      ok: false,
      problem: "port_not_allowed",
    });
  });
});

describe("addresses a webhook may connect to", () => {
  it("refuses every blocked IPv4 range at both of its edges", () => {
    for (const address of [
      "0.0.0.0",
      "0.255.255.255",
      "10.0.0.0",
      "10.255.255.255",
      "100.64.0.0",
      "100.127.255.255",
      "100.100.100.200",
      "127.0.0.1",
      "127.255.255.255",
      "169.254.0.0",
      "169.254.169.254",
      "169.254.255.255",
      "172.16.0.0",
      "172.31.255.255",
      "192.0.0.0",
      "192.0.0.192",
      "192.0.2.1",
      "192.88.99.1",
      "192.168.0.0",
      "192.168.255.255",
      "198.18.0.0",
      "198.19.255.255",
      "198.51.100.1",
      "203.0.113.1",
      "224.0.0.1",
      "239.255.255.255",
      "240.0.0.1",
      "255.255.255.255",
      "168.63.129.16",
    ])
      expect(classifyAddress(address), address).toBe("blocked");
  });

  it("allows the public addresses next to those ranges", () => {
    for (const address of [
      "1.0.0.1",
      "9.255.255.255",
      "11.0.0.0",
      "100.63.255.255",
      "100.128.0.0",
      "126.255.255.255",
      "128.0.0.0",
      "169.253.255.255",
      "169.255.0.0",
      "172.15.255.255",
      "172.32.0.0",
      "192.167.255.255",
      "192.169.0.0",
      "198.17.255.255",
      "198.20.0.0",
      "223.255.255.255",
      "168.63.129.17",
    ])
      expect(classifyAddress(address), address).toBe("public");
  });

  it("refuses IPv6 loopback, unspecified, local, multicast and tunnel forms", () => {
    for (const address of [
      "::1",
      "::",
      "0:0:0:0:0:0:0:1",
      "fe80::1",
      "febf::1",
      "fc00::1",
      "fd00:ec2::254",
      "ff02::1",
      "2001:db8::1",
      "2001::1",
      "2001:0:4136:e378:8000:63bf:3fff:fdd2",
      "2001:10::1",
      "2001:1ff::1",
      "2002:7f00:1::1",
      "2002:c0a8:101::1",
      "64:ff9b::7f00:1",
      "64:ff9b:1::1",
      "100::1",
      "3fff::1",
    ])
      expect(classifyAddress(address), address).toBe("blocked");
  });

  it("refuses IPv4 addresses written as IPv6, mapped or compatible", () => {
    for (const address of [
      "::ffff:127.0.0.1",
      "::ffff:7f00:1",
      "::ffff:169.254.169.254",
      "::ffff:a9fe:a9fe",
      "::ffff:10.0.0.1",
      "0:0:0:0:0:ffff:7f00:1",
      "::127.0.0.1",
      "::ffff:0:127.0.0.1",
      // A mapped public address is refused too: no real DNS answer is written this way.
      "::ffff:100.128.0.1",
    ])
      expect(classifyAddress(address), address).toBe("blocked");
  });

  it("allows global unicast IPv6 outside the special ranges", () => {
    for (const address of ["2003::1", "2001:200::1", "2a00:1450::1", "3fff:1000::1", "2001:db9::1"])
      expect(classifyAddress(address), address).toBe("public");
  });

  it("refuses scoped addresses and anything that is not an address", () => {
    for (const address of ["fe80::1%eth0", "2003::1%1", "", "hooks.example", "1.2.3", "::g"])
      expect(classifyAddress(address), address).toBe("blocked");
  });
});

describe("resolving a webhook host", () => {
  const answers =
    (...list: string[]): HostLookup =>
    async () =>
      list.map((address) => ({ address, family: address.includes(":") ? 6 : 4 }));

  it("answers with the first address when every answer is public", async () => {
    expect(await resolvePublicAddress("hooks.example", answers("100.128.0.1", "2003::1"))).toEqual({
      ok: true,
      address: { address: "100.128.0.1", family: 4 },
    });
  });

  it("refuses the host when any answer is blocked, wherever it is in the list", async () => {
    for (const list of [
      ["100.128.0.1", "127.0.0.1"],
      ["10.0.0.1", "100.128.0.1"],
      ["2003::1", "::ffff:127.0.0.1"],
      ["fd00:ec2::254"],
    ])
      expect(await resolvePublicAddress("hooks.example", answers(...list)), list.join()).toEqual({
        ok: false,
        problem: "blocked_address",
      });
  });

  it("reports a failed or empty lookup as a DNS failure", async () => {
    expect(await resolvePublicAddress("hooks.example", answers())).toEqual({
      ok: false,
      problem: "dns_failed",
    });
    const failing: HostLookup = async () => {
      throw Object.assign(new Error("not found"), { code: "ENOTFOUND" });
    };
    expect(await resolvePublicAddress("hooks.example", failing)).toEqual({
      ok: false,
      problem: "dns_failed",
    });
  });
});
