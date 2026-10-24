import { isIP } from "node:net";
import { MAX_WEBHOOK_URL_LENGTH } from "@prism365/contracts";

// Where a webhook may point (docs/alerts.md, ADR-0011). A webhook address is typed by an
// administrator, but it is still an address the server will connect to, so everything that
// would reach the server's own network is refused: loopback, private, link-local (and with it
// the cloud metadata services), shared address space (CGNAT), multicast, reserved and
// documentation ranges, in IPv4 and IPv6, including IPv6 forms that carry an IPv4 address.
//
// The check is made on the addresses the host resolves to at send time, and the connection then
// goes to the very address that was checked (delivery-webhook.ts), so a DNS answer that changes
// between the check and the connection cannot point the request somewhere else.

export type WebhookUrlProblem = "invalid_url" | "port_not_allowed";

export interface WebhookUrl {
  // The address without its fragment, as it is stored (encrypted) and requested.
  href: string;
  host: string;
  port: number;
  // Path and query, which carry the webhook's own token.
  path: string;
}

// Names that only ever mean the local machine or a local network.
const localSuffixes = [".localhost", ".local", ".internal", ".arpa", ".lan", ".home", ".corp"];
const label = /^(?!-)[a-z0-9-]{1,63}(?<!-)$/;

// Checks an address as typed. A host name is required (an address literal is refused, whatever
// it points to), and so is HTTPS on its default port or a port the operator listed in
// WEBHOOK_ALLOWED_PORTS. User info, whitespace, control characters and backslashes are refused
// before parsing, because the URL parser would silently drop or reinterpret them.
export function parseWebhookUrl(
  raw: string,
  allowedPorts: readonly number[],
): { ok: true; url: WebhookUrl } | { ok: false; problem: WebhookUrlProblem } {
  const invalid = { ok: false, problem: "invalid_url" } as const;
  if (raw.length === 0 || raw.length > MAX_WEBHOOK_URL_LENGTH) return invalid;
  if (/[\s\p{Cc}\p{Cf}\\]/u.test(raw)) return invalid;
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    return invalid;
  }
  if (url.protocol !== "https:" || url.username !== "" || url.password !== "") return invalid;
  // An "@" anywhere before the path would be user info to some parsers and not to others.
  const authority = raw.slice(raw.indexOf("//") + 2).split(/[/?#]/, 1)[0] ?? "";
  if (authority.includes("@")) return invalid;
  const host = url.hostname;
  if (host.startsWith("[") || isIP(host) !== 0) return invalid;
  if (!host.includes(".") || host.endsWith(".")) return invalid;
  if (host.length > 253 || !host.split(".").every((part) => label.test(part))) return invalid;
  if (host === "localhost" || localSuffixes.some((suffix) => host.endsWith(suffix))) return invalid;
  // A top-level label of digits only would make the host look like an address to some resolvers.
  if (/^[0-9]+$/.test(host.split(".").at(-1) ?? "")) return invalid;
  const port = url.port === "" ? 443 : Number(url.port);
  if (port !== 443 && !allowedPorts.includes(port))
    return { ok: false, problem: "port_not_allowed" };
  url.hash = "";
  return {
    ok: true,
    url: { href: url.href, host, port, path: `${url.pathname}${url.search}` },
  };
}

// The address's host as the page shows it after saving. Never the path or the query.
export function webhookHost(raw: string): string | null {
  try {
    return new URL(raw).hostname || null;
  } catch {
    return null;
  }
}

function ipv4Value(address: string): number | null {
  const parts = address.split(".");
  if (parts.length !== 4) return null;
  let value = 0;
  for (const part of parts) {
    if (!/^[0-9]{1,3}$/.test(part)) return null;
    const octet = Number(part);
    if (octet > 255) return null;
    value = value * 256 + octet;
  }
  return value;
}

// The eight 16-bit groups of an IPv6 address, or null. An IPv4 tail (::ffff:192.0.2.1) is
// folded into the last two groups.
export function ipv6Groups(address: string): number[] | null {
  let text = address.toLowerCase();
  const tail = text.match(/:((?:[0-9]{1,3}\.){3}[0-9]{1,3})$/);
  if (tail?.[1]) {
    const value = ipv4Value(tail[1]);
    if (value === null) return null;
    text = `${text.slice(0, -tail[1].length)}${(value >>> 16).toString(16)}:${(value & 0xffff).toString(16)}`;
  }
  const halves = text.split("::");
  if (halves.length > 2) return null;
  const parse = (part: string) => (part === "" ? [] : part.split(":"));
  const head = parse(halves[0] ?? "");
  const rest = halves.length === 2 ? parse(halves[1] ?? "") : [];
  const missing = 8 - head.length - rest.length;
  if (halves.length === 1 ? missing !== 0 : missing < 1) return null;
  const groups = [...head, ...Array<string>(halves.length === 2 ? missing : 0).fill("0"), ...rest];
  if (!groups.every((group) => /^[0-9a-f]{1,4}$/.test(group))) return null;
  return groups.map((group) => Number.parseInt(group, 16));
}

const v4 = (cidr: string): [number, number] => {
  const [base = "", bits = "32"] = cidr.split("/");
  return [ipv4Value(base) ?? 0, Number(bits)];
};

// IPv4 ranges a webhook may never reach (IANA special-purpose registry, plus Azure's platform
// address, which serves instance metadata from a public range).
export const blockedIpv4: readonly [number, number][] = [
  "0.0.0.0/8", // this network
  "10.0.0.0/8", // private
  "100.64.0.0/10", // shared address space (CGNAT), also Alibaba Cloud metadata
  "127.0.0.0/8", // loopback
  "169.254.0.0/16", // link-local, including 169.254.169.254 metadata
  "172.16.0.0/12", // private
  "192.0.0.0/24", // IETF protocol assignments, including Oracle Cloud metadata
  "192.0.2.0/24", // documentation
  "192.88.99.0/24", // 6to4 relay anycast
  "192.168.0.0/16", // private
  "198.18.0.0/15", // benchmarking, used by some proxies for fake DNS answers
  "198.51.100.0/24", // documentation
  "203.0.113.0/24", // documentation
  "224.0.0.0/4", // multicast
  "240.0.0.0/4", // reserved, including the limited broadcast address
  "168.63.129.16/32", // Azure platform address (metadata, health probes)
].map(v4);

function ipv4Blocked(value: number): boolean {
  return blockedIpv4.some(([base, bits]) => {
    const size = 2 ** (32 - bits);
    return Math.floor(value / size) === Math.floor(base / size);
  });
}

// Only global unicast (2000::/3) is allowed, which leaves out loopback, the unspecified
// address, IPv4-mapped and IPv4-compatible forms, NAT64 (64:ff9b::/96), unique local (fc00::/7,
// AWS metadata at fd00:ec2::254), link-local (fe80::/10) and multicast (ff00::/8). Inside it,
// the ranges that tunnel to or stand for something else are refused too.
function ipv6Blocked(groups: readonly number[]): boolean {
  const [first = 0, second = 0] = groups;
  if ((first & 0xe000) !== 0x2000) return true;
  // 2001::/23: IETF protocol assignments, including Teredo (2001::/32), which embeds an IPv4
  // address, and ORCHID.
  if (first === 0x2001 && second < 0x0200) return true;
  if (first === 0x2001 && second === 0x0db8) return true; // documentation
  if (first === 0x2002) return true; // 6to4, which embeds an IPv4 address
  if (first === 0x3fff && second < 0x1000) return true; // documentation (3fff::/20)
  return false;
}

// "public" when a webhook may connect to the address; anything that does not parse, or carries
// a zone index, is "blocked".
export function classifyAddress(address: string): "public" | "blocked" {
  if (address.includes("%")) return "blocked";
  const family = isIP(address);
  if (family === 4) {
    const value = ipv4Value(address);
    return value === null || ipv4Blocked(value) ? "blocked" : "public";
  }
  if (family === 6) {
    const groups = ipv6Groups(address);
    return groups === null || ipv6Blocked(groups) ? "blocked" : "public";
  }
  return "blocked";
}

export interface ResolvedAddress {
  address: string;
  family: 4 | 6;
}

// Resolves a host the way the operating system would. Every answer is checked; one refused
// answer refuses the host, so a name cannot mix a public address with a private one and hope
// the private one is picked.
export type HostLookup = (host: string) => Promise<readonly { address: string; family: number }[]>;

export async function resolvePublicAddress(
  host: string,
  lookup: HostLookup,
): Promise<
  { ok: true; address: ResolvedAddress } | { ok: false; problem: "dns_failed" | "blocked_address" }
> {
  let answers: readonly { address: string; family: number }[];
  try {
    answers = await lookup(host);
  } catch {
    return { ok: false, problem: "dns_failed" };
  }
  if (answers.length === 0) return { ok: false, problem: "dns_failed" };
  if (answers.some((answer) => classifyAddress(answer.address) !== "public"))
    return { ok: false, problem: "blocked_address" };
  const [first] = answers;
  if (!first) return { ok: false, problem: "dns_failed" };
  const family = isIP(first.address);
  return { ok: true, address: { address: first.address, family: family === 6 ? 6 : 4 } };
}
