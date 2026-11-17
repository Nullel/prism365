import { readFile } from "node:fs/promises";
import {
  type IpLocation,
  type IpLocationFailure,
  type IpLocationProvider,
  type IpLocationRequest,
  IpLocationUnavailableError,
  isIpv6Network64,
  providerNetwork,
} from "@prism365/modules";
import { z } from "zod";

// The IP location provider for IPv6 sign-ins (docs/modules/signins-audit.md §13), off unless the
// operator configures it. MaxMind's GeoIP web services, verified 2026-11-09:
// https://dev.maxmind.com/geoip/docs/web-services/requests
// https://dev.maxmind.com/geoip/docs/web-services/responses
// https://dev.maxmind.com/geoip/geolite2-free-geolocation-data
// - GET https://{host}/geoip/v2.1/city/{ip}, HTTP Basic with the account ID and license key.
//   geolite.info serves the free GeoLite City data (1000 lookups a day); geoip.maxmind.com the
//   paid GeoIP City Plus. Same response shape. HTTPS only, IPv6 supported, no batch endpoint.
// - Fields: country.iso_code, subdivisions[0].names.en, city.names.en, location.latitude,
//   location.longitude, location.accuracy_radius (km), traits.network (the record's CIDR range).
//   Empty fields are left out, not sent as null.
// - Errors carry {code, error} with a stable code: 404 IP_ADDRESS_NOT_FOUND, 400
//   IP_ADDRESS_RESERVED or IP_ADDRESS_INVALID, 401 and 403 for credentials, 402
//   INSUFFICIENT_FUNDS (for GeoLite: the daily limit is spent), 429 rate limiting, 503 retry later.
// - GeoLite data requires attribution (CC BY-SA 4.0) and its licence asks that old data be
//   destroyed within 30 days of an update. Prism365's lookup cache keeps answers forever (the
//   operator's decision, 2026-11-17; docs/modules/signins-audit.md §13 "Licence").

const services = {
  "maxmind-geolite": {
    host: "geolite.info",
    name: "MaxMind GeoLite",
    attribution:
      "This product includes GeoLite Data created by MaxMind, available from https://www.maxmind.com.",
  },
  "maxmind-geoip": {
    host: "geoip.maxmind.com",
    name: "MaxMind GeoIP City Plus",
    attribution: null,
  },
} as const;

export type IpLocationService = keyof typeof services;

export interface IpLocationConfig {
  provider: IpLocationService;
  accountId: string;
  licenseKey: string;
}

// IP_LOCATION_PROVIDER (maxmind-geolite or maxmind-geoip), IP_LOCATION_ACCOUNT_ID and
// IP_LOCATION_LICENSE_KEY_FILE. Unset provider means no lookups and nothing leaves. The key is
// read from a file only, never from the environment, the page or a request. A half-configured
// provider fails startup instead of quietly sending nothing or sending without the operator's
// full setup.
export async function ipLocationConfigFromEnv(
  env: Record<string, string | undefined>,
  read: (path: string) => Promise<string> = (path) => readFile(path, "utf8"),
): Promise<IpLocationConfig | undefined> {
  if (env.IP_LOCATION_LICENSE_KEY !== undefined)
    throw new Error(
      "IP_LOCATION_LICENSE_KEY is never read. Put the license key in a file and set IP_LOCATION_LICENSE_KEY_FILE.",
    );
  const provider = env.IP_LOCATION_PROVIDER?.trim();
  if (!provider) {
    if (env.IP_LOCATION_ACCOUNT_ID || env.IP_LOCATION_LICENSE_KEY_FILE)
      throw new Error(
        "IP location lookups need IP_LOCATION_PROVIDER as well as the account ID and key file.",
      );
    return undefined;
  }
  if (!(provider in services))
    throw new Error("IP_LOCATION_PROVIDER must be maxmind-geolite or maxmind-geoip.");
  const accountId = env.IP_LOCATION_ACCOUNT_ID?.trim() ?? "";
  if (!/^[1-9]\d{0,11}$/.test(accountId))
    throw new Error("IP_LOCATION_ACCOUNT_ID must be the MaxMind account ID, a number.");
  const path = env.IP_LOCATION_LICENSE_KEY_FILE?.trim();
  if (!path || /[\0\n\r]/.test(path))
    throw new Error("IP_LOCATION_LICENSE_KEY_FILE must name the file holding the license key.");
  let licenseKey: string;
  try {
    licenseKey = (await read(path)).trim();
  } catch {
    throw new Error("IP_LOCATION_LICENSE_KEY_FILE could not be read.");
  }
  if (!/^[\x21-\x7e]{8,128}$/.test(licenseKey))
    throw new Error("IP_LOCATION_LICENSE_KEY_FILE does not hold a license key.");
  return { provider: provider as IpLocationService, accountId, licenseKey };
}

const name = z.object({ names: z.object({ en: z.string().max(300).optional() }).optional() });
const cityResponse = z.object({
  country: z.object({ iso_code: z.string().max(8).optional() }).optional(),
  subdivisions: z.array(name).max(10).optional(),
  city: name.optional(),
  location: z
    .object({
      latitude: z.number().min(-90).max(90).optional(),
      longitude: z.number().min(-180).max(180).optional(),
      accuracy_radius: z.number().min(0).max(100_000).optional(),
    })
    .optional(),
  traits: z.object({ network: z.string().max(64).optional() }).optional(),
});
const errorResponse = z.object({ code: z.string().max(64) });

// Answers that mean "no location for this network", not a failure.
const noDataCodes: ReadonlySet<string> = new Set([
  "IP_ADDRESS_NOT_FOUND",
  "IP_ADDRESS_RESERVED",
  "IP_ADDRESS_INVALID",
]);

const MAX_BODY_BYTES = 64 * 1024;
const REQUEST_TIMEOUT_MS = 5_000;
const CONCURRENCY = 4;
// How long lookups stay paused after each kind of failure. Asking again while out of quota or
// with a refused key only adds error responses, which MaxMind answers with rate limiting.
const pauseFor: Record<IpLocationFailure, number> = {
  quota: 60 * 60_000,
  refused: 60 * 60_000,
  throttled: 5 * 60_000,
  unavailable: 60_000,
  timeout: 60_000,
  ledger: 0,
};

export type IpLocationLedger = (
  tenantId: string,
  actorId: string,
  event: Record<string, unknown>,
) => Promise<void>;

export interface MaxMindOptions {
  ledger: IpLocationLedger;
  fetchImplementation?: typeof fetch;
  now?: () => number;
  onPause?: (reason: IpLocationFailure) => void;
}

async function readBody(response: Response): Promise<string> {
  const reader = response.body?.getReader();
  if (!reader) return "";
  const chunks: Uint8Array[] = [];
  let size = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    size += value.byteLength;
    if (size > MAX_BODY_BYTES) {
      await reader.cancel().catch(() => {});
      throw new IpLocationUnavailableError("refused");
    }
    chunks.push(value);
  }
  return Buffer.concat(chunks).toString("utf8");
}

function parseJson(text: string): unknown {
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return null;
  }
}

// The network's first address in RFC 5952 form, as MaxMind recommends: "2001:db8::" for
// "2001:db8:0:0::/64". The WHATWG URL parser writes IPv6 hosts that way.
function lookupAddress(network: string): string {
  return new URL(`http://[${network.slice(0, -3)}]/`).hostname.slice(1, -1);
}

export class MaxMindIpLocation implements IpLocationProvider {
  readonly id: IpLocationService;
  readonly name: string;
  readonly attribution: string | null;
  readonly #host: string;
  readonly #authorization: string;
  readonly #ledger: IpLocationLedger;
  readonly #fetch: typeof fetch;
  readonly #now: () => number;
  readonly #onPause: (reason: IpLocationFailure) => void;
  #pausedUntil = 0;
  #pauseReason: IpLocationFailure = "unavailable";

  constructor(config: IpLocationConfig, options: MaxMindOptions) {
    const service = services[config.provider];
    this.id = config.provider;
    this.name = service.name;
    this.attribution = service.attribution;
    this.#host = service.host;
    this.#authorization = `Basic ${Buffer.from(`${config.accountId}:${config.licenseKey}`).toString("base64")}`;
    this.#ledger = options.ledger;
    this.#fetch = options.fetchImplementation ?? fetch;
    this.#now = options.now ?? Date.now;
    this.#onPause = options.onPause ?? (() => {});
  }

  #pause(reason: IpLocationFailure, milliseconds = pauseFor[reason]): never {
    if (milliseconds > 0) {
      const until = this.#now() + milliseconds;
      if (until > this.#pausedUntil) {
        this.#pausedUntil = until;
        this.#pauseReason = reason;
        this.#onPause(reason);
      }
    }
    throw new IpLocationUnavailableError(reason);
  }

  async locate(
    networks: readonly string[],
    request: IpLocationRequest,
  ): Promise<ReadonlyMap<string, IpLocation | null>> {
    // The scheduler checks this too; the provider does not rely on it.
    if (networks.length === 0 || !networks.every(isIpv6Network64))
      throw new IpLocationUnavailableError("refused");
    if (this.#now() < this.#pausedUntil) throw new IpLocationUnavailableError(this.#pauseReason);
    // Recorded before anything is sent, and nothing is sent without it. Counts only: never a
    // network, an address or a name. An application credential has no administrator to record
    // yet, so it is refused rather than sent unrecorded.
    if (!request.actorId) throw new IpLocationUnavailableError("ledger");
    try {
      await this.#ledger(request.tenantId, request.actorId, {
        type: "enrichment.ip_location",
        provider: this.id,
        module: request.moduleId,
        networks: networks.length,
        occurredAt: new Date(this.#now()).toISOString(),
      });
    } catch {
      throw new IpLocationUnavailableError("ledger");
    }

    // A run that ended while the entry was being written sends nothing. The first failure stops
    // the requests still waiting and is the answer for the whole call.
    if (request.signal.aborted) throw new IpLocationUnavailableError("timeout");
    const stop = new AbortController();
    const abort = () => stop.abort();
    request.signal.addEventListener("abort", abort, { once: true });
    const results = new Map<string, IpLocation | null>();
    const queue = [...networks];
    let failure: unknown = null;
    const worker = async () => {
      for (;;) {
        const network = queue.shift();
        if (!network || failure || stop.signal.aborted) return;
        try {
          results.set(network, await this.#lookup(network, stop.signal));
        } catch (error) {
          failure ??= error;
          stop.abort();
          return;
        }
      }
    };
    try {
      await Promise.all(Array.from({ length: Math.min(CONCURRENCY, networks.length) }, worker));
    } finally {
      request.signal.removeEventListener("abort", abort);
    }
    if (failure) {
      if (failure instanceof IpLocationUnavailableError) throw failure;
      this.#pause("unavailable");
    }
    if (request.signal.aborted) throw new IpLocationUnavailableError("timeout");
    return results;
  }

  async #lookup(network: string, signal: AbortSignal): Promise<IpLocation | null> {
    // The host is code-owned and the path is built from a checked network, so this can only
    // ever reach the configured MaxMind service.
    const url = new URL(`https://${this.#host}/geoip/v2.1/city/${lookupAddress(network)}`);
    if (url.hostname !== this.#host || url.protocol !== "https:") this.#pause("refused");
    const timeout = AbortSignal.timeout(REQUEST_TIMEOUT_MS);
    let response: Response;
    try {
      response = await this.#fetch(url, {
        method: "GET",
        headers: { Authorization: this.#authorization, Accept: "application/json" },
        redirect: "manual",
        signal: AbortSignal.any([signal, timeout]),
      });
    } catch {
      if (signal.aborted) throw new IpLocationUnavailableError("timeout");
      this.#pause(timeout.aborted ? "timeout" : "unavailable");
    }
    const body = parseJson(
      await readBody(response).catch((error: unknown) =>
        this.#pause(error instanceof IpLocationUnavailableError ? error.reason : "unavailable"),
      ),
    );
    if (response.status === 200) {
      const parsed = cityResponse.safeParse(body);
      // An answer in a shape MaxMind does not document is not cached as "no data".
      if (!parsed.success) this.#pause("refused");
      const data = parsed.data;
      const location = data.location;
      return {
        country: data.country?.iso_code?.toUpperCase() ?? null,
        state: data.subdivisions?.[0]?.names?.en ?? null,
        city: data.city?.names?.en ?? null,
        lat: location?.latitude ?? null,
        lon: location?.longitude ?? null,
        accuracyKm: location?.accuracy_radius ?? null,
        network: providerNetwork(data.traits?.network, network),
      };
    }
    const code = errorResponse.safeParse(body).data?.code ?? "";
    if ((response.status === 404 || response.status === 400) && noDataCodes.has(code)) return null;
    if (response.status === 402) this.#pause("quota");
    if (response.status === 401 || response.status === 403) this.#pause("refused");
    if (response.status === 429) {
      const seconds = Number(response.headers.get("retry-after"));
      this.#pause(
        "throttled",
        Number.isFinite(seconds) && seconds > 0
          ? Math.min(Math.max(seconds * 1_000, 60_000), 60 * 60_000)
          : pauseFor.throttled,
      );
    }
    // Redirects are never followed, and anything else is MaxMind's side.
    return this.#pause(response.status >= 300 && response.status < 400 ? "refused" : "unavailable");
  }
}

export function createIpLocationProvider(
  config: IpLocationConfig | undefined,
  options: MaxMindOptions,
): IpLocationProvider | undefined {
  return config ? new MaxMindIpLocation(config, options) : undefined;
}
