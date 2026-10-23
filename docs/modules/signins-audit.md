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

Entra Guard's `THREAT` (AbuseIPDB) and its ISP lookup (ip-api over plain HTTP) move to an
**enrichment provider** that is off by default. When enabled with the operator's consent it
adds up to 70 for abuse confidence at or above 30 and attaches ISP and organisation to the
finding. Tenant IPs do not leave the installation until that toggle is on.

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
| S0 | RED, Opus 5.5 + refuter | module spine from `MODULES.md`: manifest, registry, scheduler, cursor, `finding` and `module_run` tables with RLS helper, sink interface, PGlite-or-Postgres decision. **Implemented 2026-10-22 (TL-020); review pending.** | 2 days |
| S1 | YELLOW, spec-builder | collector, normalizer, `signins_event` and `signins_seen_*` tables, baseline updater, backfill | 1.5 days |
| S2 | YELLOW, spec-builder | rules engine, the ported and new rules, scoring, synthetic scenario tests | 1.5 days |
| S3 | YELLOW, spec-builder | the three views and the digest section | 1.5 days |
| S4 | YELLOW, spec-builder | webhook sinks and the daily digest | 0.5 day |
| S5 | RED, Fable | application credential path: certificate, application `AuditLog.Read.All`, GA consent, 24/7 polling | 1 day |
| S6 | YELLOW, spec-builder | enrichment provider with consent toggle | 0.5 day |

S1 to S4 run on the operator session, which means the first real findings at the pilot tenant arrive
while you are signed in, before S5 exists.

## 11. Before any of this touches the pilot tenant

- Entra Guard's `docker-compose.yml` has an AbuseIPDB API key in clear text, and its `.env`
  holds a client secret for some tenant. Rotate both regardless of what happens here.
- Confirm the operator account holds Reports Reader; without it the collector gets 403 and the
  module reports "unknown" by design.
