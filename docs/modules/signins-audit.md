# Module: signins.audit

**Status:** proposed 2026-10-22. Port of Entra Guard
(`~/Documents/Projects/Security & Monitoring/entra-signin-anomaly`, Python, never under git,
three rows of real data) into Prism365 as module one under `docs/MODULES.md`.

## 1. What it answers

Where do people sign in from, what is normal for each account, and which sign-ins break that
pattern. Evidence for a security conversation, not a performance signal.

## 2. Source

`GET /auditLogs/signIns?$filter=createdDateTime ge {since}&$orderby=createdDateTime asc&$top=1000`
through the existing `audit.signins.read` pack (`AuditLog.Read.All`). Interactive user
sign-ins only on v1.0, which is the right scope for location baselines. Needs Entra ID P1 on
the tenant; delegated runs also need Reports Reader, Security Reader or Global Reader on the
operator. Risk fields come back `hidden` without P2; the module treats `hidden` as absent.

Cadence (kept from Entra Guard, it was right): every 5 minutes, 10-minute lookback, 2-minute
overlap past the `last_seen_created` cursor, dedupe on sign-in id. Backfill: 30 days on first
run so baselines exist before the first alert can fire.

## 3. Normalized row (`signins_event`)

id, created_at, user_id, upn, user_display_name, app_id, app_name, resource_name,
client_app, ip, country, state, city, lat, lon, device_id, device_os, device_browser,
is_compliant, is_managed, trust_type, error_code, failure_reason, ca_status,
risk_level_signin (null when hidden), risk_level_aggregated (null when hidden), risk_state,
risk_event_types (array), is_interactive, correlation_id, collected_at.

No raw JSON by default. The IP is kept because it is security evidence; it is masked to /24 in
any view that is not an investigation of a single account.

## 4. Baselines (`signins_seen_*`, success only)

Per user, with first_seen and last_seen: ip, country, state, city, device_id, app_id,
client_app, os+browser pair. Plus `signins_user_hours`: a 24-bucket histogram of successful
sign-in hours in the tenant's configured time zone, used only for the low-severity "unusual
hour" rule and never shown as a chart of when someone works.

Baselines learn only from successes with `error_code = 0`, exactly as Entra Guard did. A
finding does not stop learning: after the operator marks it resolved, the location becomes
part of the baseline. Before that, it does not. (Entra Guard learned immediately, which meant a
compromised sign-in taught the baseline; this is the one behavioural change in the port.)

## 5. Rules

Each rule is a pure function with an id and a version. Scores add up; a finding is emitted at
40 or more, severity medium under 70, high under 90, critical at 90 or more. Thresholds and
allowlists come from the module's config, defaulted from Entra Guard's values.

Ported, with the same scores:

| id | Trigger | Score | Suppressed when |
|---|---|---|---|
| `brute_force.ip` | 5 or more failures from one IP in 15 minutes | 50 + 5 per extra failure, max 80 | never |
| `new.ip` | first success from this IP for this user | 20 | known device, or allow-listed CIDR |
| `new.country` | first success from this country for this user | 15 | known device, or allow-listed country |
| `new.state` | first success from this state for this user | 25 | known device |
| `travel.impossible` | distance / time since previous success exceeds 900 km/h within 6 hours (haversine on geoCoordinates) | 70 | known device |

New in Prism365:

| id | Trigger | Score |
|---|---|---|
| `brute_force.user` | 5 or more failures for one account in 15 minutes from any IP | 40 |
| `spray.ip` | one IP fails against 5 or more distinct accounts in 60 minutes | 60 |
| `success.after_failures` | a success for an account within 30 minutes of a `brute_force.user` or `spray.ip` window | 40 (stacks) |
| `legacy.auth` | `client_app` in the legacy set (Exchange ActiveSync, IMAP, MAPI, SMTP, POP, other clients) on a success | 30 |
| `ca.not_applied.admin` | successful sign-in by an account holding a directory role with `ca_status = notApplied` | 30 |
| `risk.entra` | `risk_level_signin` medium or high (P2 only) | 50 or 70 |
| `dormant.revival` | first success after 60 or more days without one | 25 |
| `country.outside_allowlist` | success from a country not in the tenant's expected list | 20 |
| `hour.unusual` | success in an hour bucket with under 2 % of the account's history, after 30 days of data | 10, low only, never alone |
| `signin.canary-used` | any sign-in attempt by an account marked as a canary, successful or not | critical, always (see section 12) |

Entra Guard's `THREAT` (AbuseIPDB) and its ISP lookup (ip-api over plain HTTP) move to an
**enrichment provider** that is off by default. When enabled with the operator's consent it
adds up to 70 for abuse confidence at or above 30 and attaches ISP and organisation to the
finding. Tenant IPs do not leave the installation until that toggle is on. The first enrichment
built is location for IPv6 sign-ins (section 13); reputation and ISP are not built yet.

## 6. Views

- **Sign-in map.** Successes and failures by country for 7 and 30 days as a table with bars
  (no map library; inline SVG if a map is wanted later), top failing IPs masked to /24, legacy
  auth count, Conditional Access failure count. Tenant level, no names.
- **Findings.** The shared findings list filtered to this module: severity, account, reasons,
  location, first and last seen, state, with the evidence rows expandable.
- **People (built, TL-022).** Its own tab. A list of every account that signed in during the
  last 30 days (search by name or email, or open anyone in the directory by full work email),
  and one page per account: directory details, manager, Microsoft's last sign-in times and
  licenses read live; and from stored sign-ins its places, full IP addresses, devices
  (compliance, management, join type), apps, what the baseline has learned as usual, its
  findings, and every sign-in page by page. Full addresses are shown here on purpose, because
  telling whether one account is compromised needs them. Opening a person, paging their
  sign-ins, a live directory read and an email lookup each write an access-ledger entry; the
  list of people does not. MAC addresses are not in sign-in logs; they would come from Intune.
- **History loading.** While the first 30 days load (oldest first), the view says how far it
  has reached, so an empty "last 7 days" reads as "not loaded yet", not "no sign-ins".
- **Digest section.** New findings since the last digest, by severity, plus the three
  countries and three IPs with the most failures.

### Busy windows, late records and stopped runs

Collection reads Microsoft's log in time windows of up to six hours, saving its place after
each. A request reads at most `windowRecordLimit` (500) records, about a minute at the speed
Microsoft served the first live tenant. No Graph page starts after the run's budget (135 s), so a
run always has time to save before the scheduler's 240-second limit. A window too busy to read in
one run is retried at half the size, down to five minutes, and windows grow back once they are
quick. A five-minute window that still holds more than one run can read keeps what it read, even
when the time runs out partway through a page sequence, and raises a `signins.collection_gap`
finding (operations, medium) naming the span, because a burst that dense can itself be an attack.
Only a window that had the run's whole budget may do that; one that had less is owed the whole
budget, and the next run reads that lane first. More than the cap at one millisecond is recorded as a gap
for that millisecond only; the rest of the window is still read.

Microsoft publishes some sign-ins minutes after they happen. Live collection therefore has two
saved positions:

- the **main lane** follows new sign-ins up to the run's start;
- the **settle lane** re-reads behind it. Everything before its position has been read at least
  `lateArrivalMinutes` (60) after it happened. It saves after every window, so heavy traffic or a
  long pause leaves a backlog, never a hole. When it has a backlog, the main lane stops at two
  thirds of the run after its first window; when the main lane's window takes the whole run, the
  settle lane is owed the next run and reads first, so it can fall behind by a run but never stall.

With time left over, a best-effort pass re-reads the unsettled hour newest first, so on a quiet
tenant a late record shows up on the next run rather than an hour later.

Before reading a window, the collector saves which window it is reading. If a run is stopped
mid-window (the scheduler's time limit, or a restart), the next run counts that window as failed:
it shrinks it, and after three failures of a five-minute window records an "unanswered" gap and
moves on, so nothing can stall collection. A run that stops at its budget reports more to read,
and the scheduler runs it again within seconds instead of waiting for the five-minute cadence.

### Starting, pausing and resuming

**Start collecting** is a standing choice. Collection runs on the delegated session of the
administrator who started it; a restart, a sign-out or an expired session pauses it, and it
resumes on that administrator's next signed-in page load. A session Microsoft refused a token for
while collecting (a sign-in frequency policy, a revoked refresh token) does not resume it again;
signing in again does, and so does pressing Start. Start, stop and resume run one at a time, so a
stop always wins over a resume that was already under way. **Stop collecting** ends it, and so
does another administrator starting collection. Only who chose to collect is stored
(`prism365_preference`, key `collection.intent`), never a session or token.

## 7. Notifications

Teams and Slack incoming webhooks, operator-configured, posted per finding at high or above and
as the daily digest. The Adaptive Card from Entra Guard's `notifications.py` carries over.
No mail is sent through Graph; that would need a write scope.

## 8. Retention

Events 90 days (configurable), baselines 1 year, findings indefinite. Deletion jobs and the
export already promised in `PRIVACY.md` apply to all three.

## 9. Tests

A synthetic sign-in generator that replays named scenarios against the rules: normal week,
new country on a known device (no finding), new country on an unknown device (finding), brute
force then success, impossible travel, legacy auth, spray. Collector tests use the existing
fake-fetch pattern and assert cursor advance, overlap dedupe and partial-page honesty.

## 10. Slices

| Slice | Class | Content | Est. |
|---|---|---|---|
| S0 | RED, adversarial review | module spine from `MODULES.md`: manifest, registry, scheduler, cursor, `finding` and `module_run` tables with RLS helper, sink interface, PGlite-or-Postgres decision. **Implemented 2026-10-22 (TL-020); review pending.** | 2 days |
| S1 | YELLOW | collector, normalizer, `signins_event` and `signins_seen_*` tables, baseline updater, backfill | 1.5 days |
| S2 | YELLOW | rules engine, the ported and new rules, scoring, synthetic scenario tests | 1.5 days |
| S3 | YELLOW | the three views and the digest section | 1.5 days |
| S4 | YELLOW | webhook sinks and the daily digest | 0.5 day |
| S5 | RED, Fable | application credential path: certificate, application `AuditLog.Read.All`, GA consent, 24/7 polling | 1 day |
| S6 | RED | enrichment provider with consent toggle. **IPv6 location built 2026-11-09 (TL-052, section 13); reputation and ISP not built.** | 0.5 day |

S1 to S4 run on the operator session, which means the first real findings at the pilot tenant arrive
while you are signed in, before S5 exists.

## 11. Before any of this touches the pilot tenant

- Entra Guard's `docker-compose.yml` has an AbuseIPDB API key in clear text, and its `.env`
  holds a client secret for some tenant. Rotate both regardless of what happens here.
- Confirm the operator account holds Reports Reader; without it the collector gets 403 and the
  module reports "unknown" by design.

## 12. Canary accounts

A canary is an account nobody should ever sign in to, such as a training account that was never
handed out or a mailbox kept only to be found. An administrator marks accounts as canaries, and
any sign-in attempt by one, successful or not, becomes a critical finding within one collection
run. Nothing about the rule depends on a baseline, so it also applies to the history loaded on the
first run.

### Marking an account

On a person's page, the Canary account switch marks or unmarks that account. Turning it on asks
for confirmation first and names the account; turning it off does not ask. The Sign-in audit page
has a small Canary accounts panel that lists every marked account with a button to remove it.
Opening a marked account from that list goes to its page.

### Where the list lives

One setting per tenant in `prism365_preference`, key `signins.canaries`, with no migration:

```json
{ "accounts": [{ "userId": "...", "principal": "...", "addedAt": "...", "addedBy": "..." }] }
```

At most 100 accounts. `addedBy` is the administrator's object id, as in the access ledger, never a
name. The module declares the key in its definition (`settings`), which is the only way a
collector can read a saved setting, and only inside its own namespace. The collector reads the
list once per run, so marking an account takes effect from the next run. A sign-in that was already
read before the account was marked is not looked at again. A damaged list reads as empty and a
damaged entry is skipped, so one bad row cannot stop collection or switch off the others.

`GET /v1/settings/canaries` returns the list. `POST /v1/settings/canaries` takes
`{ "add": [userId], "remove": [userId] }`, both optional but not both empty, and changes only
what it names. Marking a marked account or unmarking an unmarked one changes nothing. The list
can hold 100 accounts after the request, counting removals made in the same request.

Each added account needs the sign-in name it signs in with, so a sign-in record that carries no
user id can still be matched. The name comes from the newest stored sign-in for that account. Only
an account with no stored sign-in is looked up in the directory, which needs `User.Read.All`
(the `directory.users.read` pack) on the signed-in session. An id found in neither place refuses
the whole request and changes nothing, and so does a directory that cannot be read just now.

The access ledger gets `signins.canaries.lookup` before each read it makes (which datasets, and
how many accounts) and `signins.canaries.changed` before the list is saved (how many added, how
many removed, how many held afterwards). Counts only: never an id, a name or an address.

### The rule

`signin.canary-used` is critical, for the purpose security, about the person. A sign-in is matched
by its user id. Only a record with no user id is matched by principal name, compared after
compatibility normalisation (NFKC) and in lower case. A sign-in with a user id that is not a
canary's is never matched by name.

The dedupe key is `canary:<userId>:<UTC day>`. More attempts on the same day update one finding
and its occurrence count goes up; the evidence shown is the latest attempt, and the window covers
all of them. The next UTC day opens a new finding, so a canary that is hit again tomorrow is not
hidden behind a finding someone already resolved. A finding resolved earlier the same day reopens.

The evidence is the sign-in: when it happened, whether it succeeded or failed, the error code and
Microsoft's failure reason, the full address, the location, the app and the client. The address is
shown in full, unlike the other rules, because this finding is about one account and the address is
what an administrator would act on; the person's page shows it too.

The limitations say what the finding cannot: canaries are accounts you marked, and Prism365
cannot tell a mistaken sign-in from an attack. The sign-in log this module reads holds interactive
sign-ins, so an account used only by a background token would not appear in it.

Canary findings appear wherever sign-in findings already appear: the findings list, the person's
page and the This week page. Delivering them as a notification is a later slice.

## 13. IPv6 locations

Microsoft often places IPv6 sign-ins in the wrong city or country. Each wrong place raises
`new.country`, `new.state` and `travel.impossible` against the account's real history, so a person
on their home IPv6 connection can look like they flew abroad. When the operator configures an IP
location provider, Prism365 asks it where each successful IPv6 sign-in's network is, and uses its
answer wherever it contradicts Microsoft's. IPv4 sign-ins keep Microsoft's location; nothing about
them is sent.

### Turning it on

Off by default. The operator sets, in the environment only:

- `IP_LOCATION_PROVIDER`: `maxmind-geolite` (the free GeoLite web service, 1000 lookups a day) or
  `maxmind-geoip` (the paid GeoIP City Plus web service). Same answers in the same shape; the paid
  one is more accurate and has no daily cap.
- `IP_LOCATION_ACCOUNT_ID`: the MaxMind account ID.
- `IP_LOCATION_LICENSE_KEY_FILE`: a file holding the license key. `IP_LOCATION_LICENSE_KEY` is
  refused, so the key never sits in the environment.

Setting these is the consent decision `docs/MODULES.md` §5 asks for: from then on, tenant network
prefixes leave the installation. A half-set configuration stops startup. The API logs the provider
by name when it starts.

MaxMind was chosen over IPinfo and ip2location.io (documentation checked 2026-11-09). IPinfo's free
tier has no city or coordinates. ip2location.io gives the capital city's coordinates when it does
not know the city, which would itself cause false impossible travel. MaxMind gives an accuracy
radius and its record's own network, which the next sections rely on.

### What leaves, and what is recorded

- **The /64 network only, never the address.** One /64 is one subnet, so its location is the
  address's location, and the half of the address that identifies a device stays here. The request
  carries the network's first address (`2001:db8:1:2::`). Only global unicast IPv6 (`2000::/3`) is
  ever sent, and not 6to4 (`2002::/16`) or Teredo (`2001::/32`), whose network part carries an IPv4
  address. The scheduler checks every value at its boundary, so a module bug cannot send a full
  address, an IPv4 address or anything else.
- **Successful sign-ins only.** Failed attempts keep Microsoft's location and are never looked up,
  so nobody can spend the operator's lookups (or paid credit) by failing to sign in from addresses
  they rotate.
- **Only to the configured service.** The host (`geolite.info` or `geoip.maxmind.com`) is fixed in
  code; nothing a tenant or administrator types can change it. HTTPS only, HTTP Basic with the
  account and key, no redirects followed, 5 seconds per request, 64 KB read at most.
- **An access-ledger entry before anything is sent**, under the administrator whose collection
  made the lookup: `enrichment.ip_location` with the provider, the module and how many networks.
  Never a network or an address. If the entry cannot be written, or the run ended while it was
  being written, nothing is sent. A collection with no administrator (the future application
  credential) is refused until it has one.

### How a sign-in gets its location

Collection stores a successful IPv6 sign-in with no location source. Before the rules read it, the
same run looks up its /64, from the cache first, then from the provider, and decides:

- **The provider contradicts Microsoft**: another country, or, in the same country, a precise
  place (a city, a radius of 100 km or less) more than 25 km plus that radius from Microsoft's
  point, about the distance between two points in one city. This is the error the provider is here to correct, so its location replaces Microsoft's.
- **The provider agrees**, or is too coarse to say more than Microsoft's own country (no city, or a
  radius over 100 km). Nothing contradicts Microsoft's location, so it stands, with its state and
  coordinates. A coarse answer can never erase detail that nothing disputes.

| Outcome | Location used | `location_source` / `location_lookup` |
|---|---|---|
| The provider contradicted Microsoft | the provider's; Microsoft's kept in `ms_country`, `ms_state`, `ms_city`, `ms_lat`, `ms_lon` | provider id / `found` |
| The provider agreed, or knew only the country | Microsoft's | `microsoft` / `confirmed` |
| The provider has no data for it | Microsoft's | `microsoft` / `no_data` |
| The provider failed (outage, throttling, quota, refused key) | Microsoft's | `microsoft` / `unavailable` |
| The run's lookup time was spent first | Microsoft's | `microsoft` / `skipped` |

When the provider's location replaces Microsoft's:

- Its coordinates are kept only when it is precise. A wider answer is often a country's centre,
  which against a precise location an hour earlier reads as impossible travel. Without coordinates
  the travel rule skips the sign-in.
- Its region is compared only within one source. MaxMind names regions in English ("North
  Holland"), Microsoft often in the local language ("Noord-Holland"), so comparing the two names
  would raise the very false findings this is for. The provider's region is learned into the
  baseline under the provider's id, and `new.state` compares it with the regions that provider
  has named for the account before. When the sign-in kept coordinates, a region also counts as
  familiar when one of the account's sign-ins from another source (Microsoft, or a different
  provider) was within 100 km and its region is one the account is known for. A coarse answer
  keeps no coordinates, so its region is compared by name alone. Microsoft's names are kept as
  they are. The country (an ISO code in both) and the distance are judged as before. Switching
  provider (GeoLite to GeoIP) starts that provider's region names afresh; for sign-ins with
  coordinates, the 100 km rule keeps the account's earlier places familiar.

The travel rule measures from the account's last success that has coordinates, so a sign-in
without them does not hide the next one's travel. Dormancy (`dormant.revival`) and the finding
window are measured from the last success of any kind, as before.

Nothing waits for a later run, so a provider outage can never stop detection: the sign-in is
judged on Microsoft's location, as it would be without a provider, and the finding says so. A run
starts no new lookup once it has spent 15 seconds waiting for the provider or reached its own
deadline. One call from this module carries at most 8 networks (the scheduler accepts at most 100
from any module), four requests at a time with 5 seconds each, so a call already under way can
add about 10 seconds past that. The first failure stops lookups for the
rest of the run. After a failure the provider pauses on its own: an hour for a spent quota or a
refused key, `Retry-After` (one minute to an hour) or five minutes for throttling, a minute for an
outage. A paused provider is not asked or recorded.

Sign-ins processed before this change keep Microsoft's location; their `location_source` is null.
Baselines already learned from them stay learned. A successful IPv6 sign-in that a run cut short
left unprocessed is looked up on the next run like a new one; a failed one never is.

### Cache

`mod_signins_audit_ip_location` holds answers per tenant and provider, as the provider gave them
(with the accuracy radius), filed under the network MaxMind says its answer covers
(`traits.network`, clamped to /32 to /64), or the /64 when it does not say. Mobile networks give
every phone its own /64, so filing under MaxMind's range is what keeps a tenant inside a free daily
quota. An answer is reused for 30 days, a "no data" answer for 7, and the table's retention is 30
days. It holds networks and places, never addresses.

### What a finding says

Every IPv6 finding says where its location came from. Its evidence carries `locationSource`, the
lookup's outcome as `locationLookup`, and, when the provider's location replaced Microsoft's,
`microsoftLocation`. Its limitations say which source was used and why, and, when the provider's
location was used, how its region was compared. With GeoLite, MaxMind's required attribution
is one of the limitations. Without a provider, IPv6 findings say that Microsoft's IPv6 locations
are often wrong and that a provider can check them. Limitation text is fixed wording with
code-owned names only; place names from Microsoft or the provider appear in evidence, never in
limitations.

### Licence

GeoLite data is CC BY-SA 4.0 and needs attribution: "This product includes GeoLite Data created by
MaxMind, available from https://www.maxmind.com." Its EULA also asks that old data be destroyed
within 30 days of an update. The cache expires at 30 days. Whether a location already written into
a stored sign-in (kept 90 days) counts as old data is a question for a legal read before relying on
GeoLite in production; the paid service's terms set no such limit.

### Limits

- A location from the provider is still IP geolocation. VPNs, mobile networks and cloud proxies
  move it, and the provider can be wrong too.
- An attacker on an IPv6 network the provider cannot place is judged on Microsoft's location, as
  without a provider.
- When Microsoft is wrong within the right country and the provider knows only the country (for
  example Microsoft says Seattle for someone in Chicago), Microsoft's location stands, and so does
  any false travel it causes, as without a provider.
- When the provider contradicts Microsoft's country but is coarse, that sign-in loses its state
  and travel signals and is judged on the provider's country and the address alone. If Microsoft
  was right and the provider wrong, a real sign-in abroad can score below the finding threshold.
  This needs both sources to disagree on the country and the provider to know nothing finer.
- A lookup call already under way when collection stops (sign-out, "Stop collecting") finishes:
  at most 8 networks, already recorded in the ledger.
- Going back to a build from before this change leaves the cache table without a retention rule;
  drop `mod_signins_audit_ip_location` by hand in that case.
- In development live mode without PostgreSQL, the in-memory access ledger holds 10,000 entries;
  each lookup call adds one.
- Lookups use the delegated session's administrator for the ledger, so they run only while
  collection does.
