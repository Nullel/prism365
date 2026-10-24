import { createHmac, randomBytes, timingSafeEqual } from "node:crypto";
import { lookup as systemLookup } from "node:dns/promises";
import { Agent, request as httpsRequest, type RequestOptions } from "node:https";
import type { IncomingMessage } from "node:http";
import type { LookupFunction } from "node:net";
import type { DeliveryError } from "@prism365/contracts";
import {
  type HostLookup,
  parseWebhookUrl,
  type ResolvedAddress,
  resolvePublicAddress,
} from "./delivery-address.ts";

// Sends one webhook request (docs/alerts.md, ADR-0011). The host is resolved at send time, every
// answer is checked, and the connection goes to the checked address itself, with the host name
// kept for TLS (SNI and certificate) and the Host header, so nothing resolves the name again.
// Redirects are never followed, the whole exchange has one deadline, and at most 64 KB of the
// answer is read; nothing of it is kept. Every request is signed with the destination's own
// secret, HMAC-SHA256 over "<timestamp>.<body>", so a receiver can tell it came from here.

export const WEBHOOK_TIMEOUT_MILLISECONDS = 10_000;
export const WEBHOOK_RESPONSE_LIMIT = 65_536;

export interface WebhookTarget extends ResolvedAddress {
  host: string;
  port: number;
  path: string;
}

export interface WebhookLimits {
  timeoutMilliseconds: number;
  maxResponseBytes: number;
  signal: AbortSignal;
}

// Connects to the resolved address and answers with the status. Tests replace it; the real one is
// httpsConnector.
export type WebhookConnector = (
  target: WebhookTarget,
  request: { headers: Record<string, string>; body: string },
  limits: WebhookLimits,
) => Promise<{ status: number }>;

export const systemHostLookup: HostLookup = (host) =>
  systemLookup(host, { all: true, verbatim: true });

export type WebhookOutcome = { ok: true; status: number } | { ok: false; error: DeliveryError };

class DeadlineError extends Error {
  constructor() {
    super("The webhook did not answer in time.");
    this.name = "DeadlineError";
  }
}

// A connection error's code, mapped to a fixed category. Never the error's message.
export function webhookErrorCategory(error: unknown): DeliveryError {
  if (error instanceof DeadlineError) return "timeout";
  const name = error instanceof Error ? error.name : "";
  if (name === "AbortError" || name === "TimeoutError") return "timeout";
  const code =
    typeof error === "object" && error !== null && "code" in error ? String(error.code) : "";
  if (code === "ETIMEDOUT" || code === "ESOCKETTIMEDOUT" || code === "ABORT_ERR") return "timeout";
  if (
    code.startsWith("ERR_TLS") ||
    code.startsWith("ERR_SSL") ||
    code.includes("CERT") ||
    code === "EPROTO" ||
    code === "UNABLE_TO_VERIFY_LEAF_SIGNATURE" ||
    code === "SELF_SIGNED_CERT_IN_CHAIN" ||
    code === "DEPTH_ZERO_SELF_SIGNED_CERT"
  )
    return "tls_failed";
  return "connection_failed";
}

export function statusCategory(status: number): DeliveryError | null {
  if (status >= 200 && status < 300) return null;
  if (status >= 300 && status < 400) return "redirect";
  if (status === 408) return "timeout";
  if (status === 429) return "rate_limited";
  if (status >= 500) return "server_error";
  return "rejected";
}

// A new signing secret: 32 random bytes, base64url. The receiver uses the string itself as the
// HMAC key.
export function newSigningSecret(): string {
  return randomBytes(32).toString("base64url");
}

export function webhookSignature(secret: string, timestamp: string, body: string): string {
  return `v1=${createHmac("sha256", secret).update(`${timestamp}.${body}`).digest("hex")}`;
}

// What a receiver does with a request: recompute the signature over the timestamp and the raw
// body, compare in constant time, and refuse a timestamp more than five minutes away. Exported for
// the documentation's example and the tests.
export function verifyWebhookSignature(input: {
  secret: string;
  timestamp: string;
  body: string;
  signature: string;
  now: Date;
  toleranceSeconds?: number;
}): boolean {
  if (!/^[0-9]{1,12}$/.test(input.timestamp)) return false;
  const age = Math.abs(input.now.getTime() / 1_000 - Number(input.timestamp));
  if (age > (input.toleranceSeconds ?? 300)) return false;
  const expected = Buffer.from(webhookSignature(input.secret, input.timestamp, input.body));
  const given = Buffer.from(input.signature);
  return expected.length === given.length && timingSafeEqual(expected, given);
}

export function signedHeaders(input: {
  secret: string;
  body: string;
  deliveryId: string;
  now: Date;
}): Record<string, string> {
  const timestamp = String(Math.floor(input.now.getTime() / 1_000));
  return {
    "content-type": "application/json; charset=utf-8",
    "user-agent": "Prism365-Alerts/1",
    "x-prism365-delivery": input.deliveryId,
    "x-prism365-timestamp": timestamp,
    "x-prism365-signature": webhookSignature(input.secret, timestamp, input.body),
  };
}

export async function sendWebhook(input: {
  url: string;
  body: string;
  headers: Record<string, string>;
  allowedPorts: readonly number[];
  lookup: HostLookup;
  connect: WebhookConnector;
  timeoutMilliseconds?: number;
  maxResponseBytes?: number;
  // Ends the send early, as the deadline does: the connection is closed.
  signal?: AbortSignal;
}): Promise<WebhookOutcome> {
  // Checked again at send time: WEBHOOK_ALLOWED_PORTS may have changed since it was saved.
  const parsed = parseWebhookUrl(input.url, input.allowedPorts);
  if (!parsed.ok) return { ok: false, error: parsed.problem };
  const { url } = parsed;
  const timeout = input.timeoutMilliseconds ?? WEBHOOK_TIMEOUT_MILLISECONDS;
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  let stop: () => void = () => {};
  const deadline = new Promise<never>((_resolve, reject) => {
    stop = () => {
      controller.abort(new DeadlineError());
      reject(new DeadlineError());
    };
    timer = setTimeout(stop, timeout);
  });
  deadline.catch(() => {});
  if (input.signal?.aborted) stop();
  input.signal?.addEventListener("abort", stop, { once: true });
  try {
    const resolved = await Promise.race([resolvePublicAddress(url.host, input.lookup), deadline]);
    if (!resolved.ok) return { ok: false, error: resolved.problem };
    const answer = await Promise.race([
      input.connect(
        { ...resolved.address, host: url.host, port: url.port, path: url.path },
        { headers: input.headers, body: input.body },
        {
          timeoutMilliseconds: timeout,
          maxResponseBytes: input.maxResponseBytes ?? WEBHOOK_RESPONSE_LIMIT,
          signal: controller.signal,
        },
      ),
      deadline,
    ]);
    const problem = statusCategory(answer.status);
    return problem ? { ok: false, error: problem } : { ok: true, status: answer.status };
  } catch (error) {
    return { ok: false, error: webhookErrorCategory(error) };
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    input.signal?.removeEventListener("abort", stop);
    controller.abort();
  }
}

// Reads and drops at most `limit` bytes of an answer, then closes it. Only the status matters.
function drain(response: IncomingMessage, limit: number): void {
  let read = 0;
  response.on("data", (chunk: Buffer) => {
    read += chunk.length;
    if (read > limit) response.destroy();
  });
  response.on("error", () => {});
  response.resume();
}

// The real connector. `request` is replaceable so the tests can check the options it sends and
// how it reads an answer without opening a socket; production always uses https.request, with
// certificate checks on.
export function httpsConnector(request: typeof httpsRequest = httpsRequest): WebhookConnector {
  return (target, outgoing, limits) =>
    new Promise((resolve, reject) => {
      // The address was resolved and checked already. A connection that asks to resolve anything
      // gets that same address and nothing else.
      const pinned: LookupFunction = (_hostname, options, callback) => {
        if (options.all) callback(null, [{ address: target.address, family: target.family }]);
        else callback(null, target.address, target.family);
      };
      const options: RequestOptions = {
        host: target.address,
        family: target.family,
        port: target.port,
        path: target.path,
        method: "POST",
        servername: target.host,
        lookup: pinned,
        headers: {
          ...outgoing.headers,
          host: target.port === 443 ? target.host : `${target.host}:${target.port}`,
          "content-length": String(Buffer.byteLength(outgoing.body)),
        },
        // A fresh agent per request: no pooled socket from an earlier address, and no proxy
        // taken from the environment.
        agent: new Agent({ keepAlive: false }),
        rejectUnauthorized: true,
        minVersion: "TLSv1.2",
        timeout: limits.timeoutMilliseconds,
        signal: limits.signal,
      };
      const outgoingRequest = request(options, (response) => {
        drain(response, limits.maxResponseBytes);
        resolve({ status: response.statusCode ?? 0 });
      });
      outgoingRequest.on("error", (error: Error) => reject(error));
      outgoingRequest.on("timeout", () => outgoingRequest.destroy(new DeadlineError()));
      outgoingRequest.end(outgoing.body);
    });
}
