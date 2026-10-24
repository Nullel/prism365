import { EventEmitter } from "node:events";
import type { IncomingMessage } from "node:http";
import type { RequestOptions } from "node:https";
import { PassThrough } from "node:stream";
import { describe, expect, it } from "vitest";
import type { HostLookup } from "../src/delivery-address.ts";
import {
  httpsConnector,
  newSigningSecret,
  sendWebhook,
  signedHeaders,
  statusCategory,
  verifyWebhookSignature,
  type WebhookConnector,
  type WebhookTarget,
} from "../src/delivery-webhook.ts";

// Sending one webhook: resolve once, check, connect to the checked address, no redirects, one
// deadline, a bounded read, and a signature the receiver can check. Everything runs against fakes;
// no test opens a socket.

const url = "https://hooks.example/services/synthetic?sig=token";
const body = JSON.stringify({ text: "synthetic" });

function recorder(status = 200) {
  const targets: WebhookTarget[] = [];
  const connect: WebhookConnector = async (target) => {
    targets.push(target);
    return { status };
  };
  return { targets, connect };
}

const fixed =
  (...list: string[]): HostLookup =>
  async () =>
    list.map((address) => ({ address, family: address.includes(":") ? 6 : 4 }));

const send = (overrides: Partial<Parameters<typeof sendWebhook>[0]> = {}) =>
  sendWebhook({
    url,
    body,
    headers: { "content-type": "application/json" },
    allowedPorts: [],
    lookup: fixed("100.128.0.1"),
    connect: recorder().connect,
    ...overrides,
  });

describe("sending a webhook", () => {
  it("connects to the address it checked, keeping the host name for TLS", async () => {
    const { targets, connect } = recorder();
    expect(await send({ connect })).toEqual({ ok: true, status: 200 });
    expect(targets).toEqual([
      {
        address: "100.128.0.1",
        family: 4,
        host: "hooks.example",
        port: 443,
        path: "/services/synthetic?sig=token",
      },
    ]);
  });

  it("resolves once, so a DNS answer that changes after the check is never used", async () => {
    const answers = ["100.128.0.1", "127.0.0.1"];
    let lookups = 0;
    const rebinding: HostLookup = async () => {
      const address = answers[Math.min(lookups, answers.length - 1)] ?? "127.0.0.1";
      lookups += 1;
      return [{ address, family: 4 }];
    };
    const { targets, connect } = recorder();
    expect(await send({ lookup: rebinding, connect })).toEqual({ ok: true, status: 200 });
    expect(lookups).toBe(1);
    expect(targets.map((target) => target.address)).toEqual(["100.128.0.1"]);
    // The next delivery resolves again and is refused before connecting.
    expect(await send({ lookup: rebinding, connect })).toEqual({
      ok: false,
      error: "blocked_address",
    });
    expect(targets).toHaveLength(1);
  });

  it("never connects when the host resolves to a blocked address", async () => {
    for (const address of ["169.254.169.254", "10.1.2.3", "::1", "::ffff:127.0.0.1", "fd00::1"]) {
      const { targets, connect } = recorder();
      expect(await send({ lookup: fixed(address), connect }), address).toEqual({
        ok: false,
        error: "blocked_address",
      });
      expect(targets).toEqual([]);
    }
  });

  it("checks the address and its port again at send time", async () => {
    const { targets, connect } = recorder();
    expect(await send({ url: "https://hooks.example:8443/x", connect })).toEqual({
      ok: false,
      error: "port_not_allowed",
    });
    expect(await send({ url: "https://user@hooks.example/x", connect })).toEqual({
      ok: false,
      error: "invalid_url",
    });
    expect(targets).toEqual([]);
    expect(
      await send({ url: "https://hooks.example:8443/x", allowedPorts: [8443], connect }),
    ).toEqual({ ok: true, status: 200 });
    expect(targets[0]?.port).toBe(8443);
  });

  it("treats a redirect as a failure and never follows it", async () => {
    for (const status of [301, 302, 303, 307, 308]) {
      const { targets, connect } = recorder(status);
      expect(await send({ connect }), String(status)).toEqual({ ok: false, error: "redirect" });
      expect(targets).toHaveLength(1);
    }
  });

  it("maps answers to fixed categories", () => {
    expect(statusCategory(200)).toBeNull();
    expect(statusCategory(202)).toBeNull();
    expect(statusCategory(400)).toBe("rejected");
    expect(statusCategory(404)).toBe("rejected");
    expect(statusCategory(408)).toBe("timeout");
    expect(statusCategory(429)).toBe("rate_limited");
    expect(statusCategory(503)).toBe("server_error");
  });

  it("gives up at the deadline, whether DNS or the endpoint is stuck", async () => {
    const never = () => new Promise<never>(() => {});
    const started = Date.now();
    expect(await send({ lookup: never, timeoutMilliseconds: 50 })).toEqual({
      ok: false,
      error: "timeout",
    });
    let signal: AbortSignal | null = null;
    const stuck: WebhookConnector = (_target, _request, limits) => {
      signal = limits.signal;
      return never();
    };
    expect(await send({ connect: stuck, timeoutMilliseconds: 50 })).toEqual({
      ok: false,
      error: "timeout",
    });
    // The connection is told to stop, not left running.
    expect((signal as AbortSignal | null)?.aborted).toBe(true);
    expect(Date.now() - started).toBeLessThan(2_000);
  });

  it("reports connection and certificate failures by category", async () => {
    const failing =
      (code: string): WebhookConnector =>
      async () => {
        throw Object.assign(new Error("synthetic"), { code });
      };
    expect(await send({ connect: failing("ECONNREFUSED") })).toEqual({
      ok: false,
      error: "connection_failed",
    });
    expect(await send({ connect: failing("CERT_HAS_EXPIRED") })).toEqual({
      ok: false,
      error: "tls_failed",
    });
    expect(await send({ connect: failing("ERR_TLS_CERT_ALTNAME_INVALID") })).toEqual({
      ok: false,
      error: "tls_failed",
    });
  });
});

describe("the HTTPS connector", () => {
  // A stand-in for https.request that records the options and answers with `chunks`.
  function fakeRequest(status: number, chunks: Buffer[]) {
    const seen: { options: RequestOptions | null; body: string; destroyed: boolean } = {
      options: null,
      body: "",
      destroyed: false,
    };
    const request = ((options: RequestOptions, callback: (response: IncomingMessage) => void) => {
      seen.options = options;
      const outgoing = new EventEmitter() as EventEmitter & {
        end(body: string): void;
        destroy(): void;
      };
      outgoing.end = (sent: string) => {
        seen.body = sent;
        const response = new PassThrough() as unknown as IncomingMessage & PassThrough;
        response.statusCode = status;
        // Cut off: destroyed before the whole answer was written (a stream that ended destroys
        // itself, which does not count).
        let written = false;
        const originalDestroy = response.destroy.bind(response);
        response.destroy = ((error?: Error) => {
          if (!written) seen.destroyed = true;
          return originalDestroy(error);
        }) as typeof response.destroy;
        callback(response);
        for (const chunk of chunks) if (!seen.destroyed) response.write(chunk);
        written = true;
        if (!seen.destroyed) response.end();
      };
      outgoing.destroy = () => {};
      return outgoing;
    }) as unknown as typeof import("node:https").request;
    return { request, seen };
  }

  const target: WebhookTarget = {
    address: "100.128.0.1",
    family: 4,
    host: "hooks.example",
    port: 443,
    path: "/services/synthetic?sig=token",
  };
  const limits = () => ({
    timeoutMilliseconds: 1_000,
    maxResponseBytes: 65_536,
    signal: new AbortController().signal,
  });

  it("connects to the address, names the host for TLS, verifies certificates, uses no pool", async () => {
    const { request, seen } = fakeRequest(204, []);
    const answer = await httpsConnector(request)(target, { headers: { a: "b" }, body }, limits());
    expect(answer).toEqual({ status: 204 });
    const options = seen.options as RequestOptions;
    expect(options.host).toBe("100.128.0.1");
    expect(options.servername).toBe("hooks.example");
    expect(options.rejectUnauthorized).toBe(true);
    expect(options.method).toBe("POST");
    expect(options.path).toBe("/services/synthetic?sig=token");
    expect(options.headers).toMatchObject({ host: "hooks.example", a: "b" });
    expect((options.agent as { keepAlive?: boolean } | undefined)?.keepAlive).toBe(false);
    expect(seen.body).toBe(body);
    // Asked to resolve anything, the connection gets the checked address and nothing else.
    const resolved = await new Promise((resolve) =>
      (options.lookup as (h: string, o: object, c: (...args: unknown[]) => void) => void)(
        "elsewhere.example",
        {},
        (_error, address) => resolve(address),
      ),
    );
    expect(resolved).toBe("100.128.0.1");
    const all = await new Promise((resolve) =>
      (options.lookup as (h: string, o: object, c: (...args: unknown[]) => void) => void)(
        "elsewhere.example",
        { all: true },
        (_error, addresses) => resolve(addresses),
      ),
    );
    expect(all).toEqual([{ address: "100.128.0.1", family: 4 }]);
  });

  it("puts a non-default port in the Host header", async () => {
    const { request, seen } = fakeRequest(200, []);
    await httpsConnector(request)({ ...target, port: 8443 }, { headers: {}, body }, limits());
    expect(seen.options?.headers).toMatchObject({ host: "hooks.example:8443" });
  });

  it("stops reading an answer after 64 KB", async () => {
    const { request, seen } = fakeRequest(
      200,
      Array.from({ length: 40 }, () => Buffer.alloc(16_384, 120)),
    );
    expect(await httpsConnector(request)(target, { headers: {}, body }, limits())).toEqual({
      status: 200,
    });
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(seen.destroyed).toBe(true);
  });

  it("reads a small answer to the end without cutting it off", async () => {
    const { request, seen } = fakeRequest(200, [Buffer.from("ok")]);
    await httpsConnector(request)(target, { headers: {}, body }, limits());
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(seen.destroyed).toBe(false);
  });
});

describe("request signatures", () => {
  const now = new Date("2026-10-24T12:00:00.000Z");

  it("signs the timestamp and body so the receiver can verify them", () => {
    const secret = newSigningSecret();
    expect(secret).toMatch(/^[A-Za-z0-9_-]{43}$/);
    const headers = signedHeaders({ secret, body, deliveryId: "delivery-1", now });
    expect(headers["x-prism365-timestamp"]).toBe("1792843200");
    expect(headers["x-prism365-signature"]).toMatch(/^v1=[0-9a-f]{64}$/);
    expect(headers["x-prism365-delivery"]).toBe("delivery-1");
    const check = (overrides: Partial<Parameters<typeof verifyWebhookSignature>[0]> = {}) =>
      verifyWebhookSignature({
        secret,
        timestamp: headers["x-prism365-timestamp"] ?? "",
        body,
        signature: headers["x-prism365-signature"] ?? "",
        now,
        ...overrides,
      });
    expect(check()).toBe(true);
    expect(check({ body: `${body} ` })).toBe(false);
    expect(check({ secret: newSigningSecret() })).toBe(false);
    expect(check({ timestamp: "1792843201" })).toBe(false);
    expect(check({ now: new Date(now.getTime() + 301_000) })).toBe(false);
    expect(check({ signature: "v1=00" })).toBe(false);
  });

  it("makes a different secret every time", () => {
    expect(new Set(Array.from({ length: 20 }, newSigningSecret)).size).toBe(20);
  });
});
