// Preflight for the IPv6 location pilot (docs/modules/signins-audit.md §13, "Piloting it").
// Reads IP_LOCATION_PROVIDER, IP_LOCATION_ACCOUNT_ID and IP_LOCATION_LICENSE_KEY_FILE exactly as
// the API does, then asks MaxMind about one fixed public network, Google Public DNS's
// 2001:4860:4860:0::/64, through the same client the API uses. It takes no network argument, so it
// can never send tenant data. It spends one lookup of the account's quota and stores nothing.
//
//   IP_LOCATION_PROVIDER=maxmind-geolite IP_LOCATION_ACCOUNT_ID=123456 \
//   IP_LOCATION_LICENSE_KEY_FILE=./secrets/maxmind-license-key \
//   npx tsx scripts/pilot/ip-location-check.ts
//
// Exit 0: the key works and the answer parsed. 1: MaxMind refused or failed. 2: configuration.
import { IpLocationUnavailableError, ipv6Network64, isIpv6Network64 } from "@prism365/modules";
import { ipLocationConfigFromEnv, MaxMindIpLocation } from "../../apps/api/src/ip-location.ts";

// Built by the same function the collector uses, so it is in the one form the client accepts
// ("2001:4860:4860:0::/64"); a hand-written "::/64" is refused before anything is sent.
const PUBLIC_NETWORK = ipv6Network64("2001:4860:4860::8888");

const reasons: Record<string, string> = {
  quota: "the daily lookups (GeoLite) or the credit (GeoIP) are spent",
  refused: "MaxMind refused the account ID or key, or this provider is not enabled for the account",
  throttled: "MaxMind is rate limiting this account",
  unavailable: "MaxMind could not be reached or answered with an error",
  timeout: "MaxMind did not answer in time",
  ledger: "the ledger refused the call (not expected here)",
};

if (process.argv.length > 2) {
  console.error("This check takes no arguments; it only ever looks up a fixed public network.");
  process.exit(2);
}
if (!isIpv6Network64(PUBLIC_NETWORK)) {
  console.error("The check's own network is not in the form the client accepts. Nothing was sent.");
  process.exit(2);
}

let config: Awaited<ReturnType<typeof ipLocationConfigFromEnv>>;
try {
  config = await ipLocationConfigFromEnv(process.env);
} catch (error) {
  console.error(error instanceof Error ? error.message : "The configuration could not be read.");
  process.exit(2);
}
if (!config) {
  console.error("IP_LOCATION_PROVIDER is not set, so the API would send nothing. Set it to try.");
  process.exit(2);
}

// No tenant and no administrator: this network is public and not a sign-in, so the API's
// access-ledger entry has nothing to record. The client still requires one, so it gets a no-op.
// The HTTP status is printed so a failure shows whether MaxMind answered at all; never a header
// or a body.
let status: number | null = null;
const provider = new MaxMindIpLocation(config, {
  ledger: async () => {},
  fetchImplementation: async (input, init) => {
    const response = await fetch(input, init);
    status = response.status;
    return response;
  },
});
console.log(`Provider: ${provider.name} (${config.provider}), account ${config.accountId}`);
console.log(`Looking up ${PUBLIC_NETWORK} (Google Public DNS, not tenant data).`);
try {
  const answers = await provider.locate([PUBLIC_NETWORK], {
    tenantId: "00000000-0000-0000-0000-000000000000",
    actorId: "preflight",
    moduleId: "pilot-preflight",
    signal: AbortSignal.timeout(15_000),
  });
  const answer = answers.get(PUBLIC_NETWORK);
  if (!answer) {
    console.log("MaxMind accepted the key but has no location for this network.");
  } else {
    const place = [answer.city, answer.state, answer.country].filter(Boolean).join(", ");
    console.log(`Answer: ${place || "no place named"}`);
    console.log(
      `Point: ${answer.lat ?? "none"}, ${answer.lon ?? "none"}; accuracy ${answer.accuracyKm ?? "unknown"} km; filed under ${answer.network}`,
    );
  }
  if (provider.attribution) console.log(`Attribution shown in findings: ${provider.attribution}`);
  console.log("OK: the API will be able to look up IPv6 sign-ins with this configuration.");
} catch (error) {
  const reason = error instanceof IpLocationUnavailableError ? error.reason : "unavailable";
  const answered = status === null ? "no HTTP answer" : `MaxMind answered HTTP ${status}`;
  console.error(`Failed (${reason}, ${answered}): ${reasons[reason] ?? reason}.`);
  process.exit(1);
}
