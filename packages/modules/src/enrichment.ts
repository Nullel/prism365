import { isIP } from "node:net";

// Enrichment from outside Microsoft (docs/MODULES.md §5). A module asks for a kind in its
// definition, like a read pack; the operator switches a provider on; only then does the
// collector's context offer it. What may leave the installation is checked here, at the
// scheduler boundary, so a module cannot send more than the kind allows even by mistake.

export type EnrichmentKind = "ip.location";

export const enrichmentKinds: readonly EnrichmentKind[] = ["ip.location"];

// What a provider reports for one network. `country` is ISO 3166-1 alpha-2, upper case.
// `accuracyKm` is the provider's own radius around the coordinates within which the address
// probably lies, or null when it gives none. `network` is the provider's own record range that
// holds the asked network, in the form `ipv6Prefix` returns, when the provider reports one: every
// address in it gets the same answer, so a module may reuse the answer for all of it.
export interface IpLocation {
  country: string | null;
  state: string | null;
  city: string | null;
  lat: number | null;
  lon: number | null;
  accuracyKm: number | null;
  network: string | null;
}

export interface IpLocationRequest {
  tenantId: string;
  // The administrator whose collection asked, for the access ledger. Null for an application
  // credential.
  actorId: string | null;
  moduleId: string;
  signal: AbortSignal;
}

// Implemented by the host (apps/api) when the operator configured a provider. `locate` receives
// IPv6 /64 networks in the form `ipv6Network64` returns and answers each with a location, or
// null when the provider has none for it. It writes the access-ledger entry before anything is
// sent and throws IpLocationUnavailableError when it cannot answer now.
export interface IpLocationProvider {
  // Short, code-owned id stored with each result, such as "maxmind".
  readonly id: string;
  // Name for limitations and evidence text, such as "MaxMind GeoLite".
  readonly name: string;
  // Credit the provider's data licence requires wherever its answers are shown, or null.
  readonly attribution: string | null;
  locate(
    networks: readonly string[],
    request: IpLocationRequest,
  ): Promise<ReadonlyMap<string, IpLocation | null>>;
}

export type IpLocationFailure =
  | "throttled"
  | "quota"
  | "unavailable"
  | "timeout"
  | "refused"
  | "ledger";

// A fixed reason, never provider text, so it is safe to log and store.
export class IpLocationUnavailableError extends Error {
  readonly reason: IpLocationFailure;
  constructor(reason: IpLocationFailure) {
    super(`The IP location provider could not answer (${reason}).`);
    this.name = "IpLocationUnavailableError";
    this.reason = reason;
  }
}

// The most networks one `locate` call may carry.
export const IP_LOCATION_BATCH_LIMIT = 100;

function ipv6Groups(address: string): number[] | null {
  if (isIP(address) !== 6 || address.includes("%") || address.includes(".")) return null;
  const halves = address.split("::");
  if (halves.length > 2) return null;
  const parse = (value: string | undefined) => (value ? value.split(":") : []);
  const left = parse(halves[0]);
  const right = halves.length === 2 ? parse(halves[1]) : [];
  const missing = 8 - left.length - right.length;
  if (halves.length === 1 ? missing !== 0 : missing < 1) return null;
  const groups = [...left, ...Array.from({ length: missing }, () => "0"), ...right];
  const values = groups.map((group) =>
    /^[0-9a-f]{1,4}$/i.test(group) ? Number.parseInt(group, 16) : Number.NaN,
  );
  return values.length === 8 && values.every((value) => Number.isInteger(value)) ? values : null;
}

const hex = (value: number) => value.toString(16);

const globalUnicast = (groups: readonly number[]) => ((groups[0] ?? 0) & 0xe000) === 0x2000;

function prefixOf(groups: readonly number[], bits: number): string {
  const count = Math.ceil(bits / 16);
  const kept = groups.slice(0, count).map((group, index) => {
    const left = bits - index * 16;
    return left >= 16 ? group : group & ((0xffff << (16 - left)) & 0xffff);
  });
  return `${kept.map(hex).join(":")}::/${bits}`;
}

// The network of `bits` (1 to 64) that a global unicast IPv6 address or network belongs to, as
// "2001:db8::/32": the first ceil(bits / 16) groups, masked, in lower-case hex without leading
// zeros, then "::/bits". Anything that is not global unicast (2000::/3) is null: IPv4, IPv4
// written as IPv6, link-local, unique local, multicast, scoped.
export function ipv6Prefix(address: string | null, bits: number): string | null {
  if (!address || !Number.isInteger(bits) || bits < 1 || bits > 64) return null;
  const groups = ipv6Groups(address.trim());
  return groups && globalUnicast(groups) ? prefixOf(groups, bits) : null;
}

// The /64 network an IPv6 address belongs to, as "2001:db8:0:1::/64". One /64 is one subnet, so
// its location is the address's location, and the part that identifies a device never leaves
// the installation. 6to4 (2002::/16) and Teredo (2001::/32) are null: their network part carries
// an IPv4 address, which would leave with it.
export function ipv6Network64(address: string | null): string | null {
  const network = ipv6Prefix(address, 64);
  if (!network) return null;
  const [first = "", second = ""] = network.split(":");
  if (first === "2002" || (first === "2001" && second === "0")) return null;
  return network;
}

// A provider's CIDR range ("2001:db8::/32") as the network that holds `network64` at the
// provider's prefix length, clamped to /32 at the shortest, or null when the range does not hold
// it. A range longer than /64 is answered for its /64. The result is what `ipv6Prefix` returns.
export function providerNetwork(cidr: unknown, network64: string): string | null {
  if (typeof cidr !== "string" || !isIpv6Network64(network64)) return null;
  const match = /^([0-9a-f:]+)\/(\d{1,3})$/i.exec(cidr.trim());
  if (!match?.[1] || !match[2]) return null;
  const bits = Number(match[2]);
  if (!Number.isInteger(bits) || bits < 1 || bits > 128) return null;
  const base = network64.slice(0, -3);
  const compared = Math.min(64, bits);
  const range = ipv6Prefix(match[1], compared);
  if (!range || range !== ipv6Prefix(base, compared)) return null;
  return ipv6Prefix(base, Math.min(64, Math.max(32, bits)));
}

// The networks from /64 down to /32 that hold `network64`, longest first: where a cached answer
// for it could be filed.
export function containingNetworks(network64: string): string[] {
  if (!isIpv6Network64(network64)) return [];
  const base = network64.slice(0, -3);
  const networks: string[] = [];
  for (let bits = 64; bits >= 32; bits -= 1) {
    const network = ipv6Prefix(base, bits);
    if (network) networks.push(network);
  }
  return networks;
}

const networkPattern = /^(?:(?:0|[1-9a-f][0-9a-f]{0,3}):){4}:\/64$/;

// True only for a string `ipv6Network64` could have returned.
export function isIpv6Network64(value: unknown): value is string {
  return (
    typeof value === "string" &&
    networkPattern.test(value) &&
    ipv6Network64(value.slice(0, -3)) === value
  );
}
