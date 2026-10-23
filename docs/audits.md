# Audits

Status: built (TL-025). Packages: `packages/audits` (the audits and their run store),
`apps/api/src/audits.ts` (the runner), `apps/web/src/Audits.tsx` (the page). Decision: ADR-0008.

An audit is a check an administrator runs when they need it. It reads Microsoft once, with the
signed-in administrator's delegated session, and keeps the result so it can be reopened, compared
with earlier runs and exported. An audit can also be set to run daily or weekly while collection
is on (section 12), and each audit's summary numbers are kept as a trend (section 11). No audit
changes the tenant.

## 1. The built-in audits

| Audit | Question it answers | Needs |
|---|---|---|
| Inactive accounts (`accounts.inactive`) | Which accounts have not signed in, and for how long? Views for 7, 30, 60, 90, 120, 180 and 365+ days, never signed in, and attempts with no success. | `User.Read.All`, `AuditLog.Read.All`, Microsoft Entra ID P1 or P2, a role such as Reports Reader, Security Reader or Global Reader |
| Licenses to reclaim (`licenses.reclaim`) | Which licensed accounts are disabled or 30+ days inactive, and was each license assigned directly or by a group? | `User.Read.All`; sign-in dates as above, or it lists disabled accounts only |
| MFA registration (`mfa.registration`) | Who has no strong method registered, which admins lack MFA, who relies on text or calls? | `AuditLog.Read.All`, a role such as Reports Reader, Security Reader, Security Administrator or Global Reader |
| Guest accounts (`guests.review`) | Which invitations were never accepted, which guests never signed in or went quiet, from which domains? | `User.Read.All`; sign-in dates as above, or invitation states only |
| License capacity (`licenses.capacity`) | Which products have no seats left, more assigned than owned, an expired or suspended subscription, or unassigned seats? | `LicenseAssignment.Read.All`, a role such as Directory Readers, Global Reader or License Administrator |
| Conditional Access policies (`identity.conditional-access`) | Which policies exist and what do they require and leave out? Does any enabled policy block legacy authentication or ask administrators for MFA? | `Policy.Read.All`, a role such as Global Reader, Security Reader, Security Administrator or Conditional Access Administrator, Microsoft Entra ID P1; `User.Read.All` and `LicenseAssignment.Read.All` add detail |
| Risky users (`identity.risky-users`) | Which accounts does Microsoft Entra ID Protection mark as at risk or compromised, how high, and has the risk been dealt with? | `IdentityRiskyUser.Read.All`, a role such as Global Reader, Security Reader, Security Operator or Security Administrator, Microsoft Entra ID P2 |

`LicenseAssignment.Read.All`, when granted, also gives the account audits license names instead
of Microsoft's product IDs.

## 2. How inactivity is counted

From Microsoft's `signInActivity` on each user (Graph v1.0, checked 2026-10-22):

1. **The last successful sign-in**, interactive or background (`lastSuccessfulSignInDateTime`).
   This is when the account was actually used.
2. Without one, **the last attempt**, interactive or background, successful or not. Microsoft
   has recorded successful sign-ins only since 1 December 2023 and did not backfill them, so an
   account whose last success was earlier has no success on record. Counting from the attempt
   can only make an account look more recent than it is, never older, so nothing is flagged on
   a guess. An account with attempts after December 2023 and no success on record is also tagged
   "Attempts, no success": a broken account, or someone guessing its password.
3. With no sign-in on record at all, **the account's creation**, or April 2020 if it is older,
   because Microsoft's interactive records start then. These accounts are "Never signed in".

A date Microsoft could not have recorded is **unreadable**: malformed, before its records start
(April 2020, or November 2023 for a successful sign-in), or in the future. Counting from a value
such as `0001-01-01` would make an active account look years inactive, so such an account gets
"Unreadable dates", no day count and no range, and the run is marked partial.

The "Counted from" column says which of the three applies to each row. Ranges are cumulative: an
account inactive 130 days is in 7+, 30+, 60+, 90+ and 120+.

## 3. Using the page

**Audits** in the sidebar, under Audit, lists every audit with its last result. **Run audit**
starts one; the page polls until it finishes, typically seconds for a small tenant. One run per
audit at a time: pressing Run again while it runs joins the run in progress.

An audit's page has summary tiles (each opens its view), preset views, a filter for each
fixed-value column, search, sortable columns, a column chooser, and **Export CSV**, which exports
exactly the rows and order on screen. Names open the person's page. The Run menu reopens any of
the last 10 results.

A run that cannot finish says why in plain words: Microsoft refused (with the permission, role or
license it needs, and Microsoft's request ID), throttling, a slow answer, a session that ended, or
the 10-minute limit. A run Microsoft answered only in part (an unreadable record, the
10,000-account read limit, sign-in dates refused) is marked **partial** and its notes say what is
missing.

## 4. Setup

The audits use the read packs the rest of Prism365 uses. With the sign-in audit and People set
up (`directory.users.read`, `user.licenses.read`, `audit.signins.read`), every audit is
available. The MFA audit reads `reports/authenticationMethods/userRegistrationDetails` through
its own pack, `reports.authentication-methods.read`, which uses the same `AuditLog.Read.All`; an
installation that enabled `audit.signins.read` needs nothing more. The identity audits
use two more packs: `identity.conditional-access.read` (`Policy.Read.All`) and
`identity.risk.read` (`IdentityRiskyUser.Read.All`); the Conditional Access audit also uses
`directory.users.read` and `tenant.licenses.read` when they are granted. A pack's permission must be
both enabled in `ENABLED_READ_FEATURES` and granted by the session. The page says which is
missing and links to **Grant access**.

## 5. Storage and retention

Runs live in `prism365_audit_run` (core migration `004-audit-runs.sql`), a tenant table with
forced row-level security like every other. Each row holds who started the run, when, how it
ended, its summary and notes, and its result rows: names, sign-in names, last sign-in dates,
account state and license names. Per audit, the newest 10 successful runs and the newest 5 failed
runs are kept, counted apart so a string of failures can never push out the last good result;
older ones are deleted when a run finishes. A run left "running" by a stopped server is closed as
interrupted the next time the audits are read. A run also stores the scope it used (`run_scope`)
and, for the account audits, the domains it read before that scope (`domain_counts`). A stopped
run also stops reading Microsoft: the
cancel reaches the request in flight. Text Microsoft returns is cleaned before it is stored (NUL
characters and broken UTF-16 removed), so one odd value cannot make a result impossible to save.

Starting a run, opening a finished result and exporting it each write an access-ledger entry
(`audit.run.started`, `audit.run.viewed`, `audit.run.exported`) naming the audit and run, never
an account. The start entry is written once per run, before anything is read, and lists every
permission the run reads with, optional ones included. Exported files are outside Prism365's retention once downloaded.

## 6. API

All routes need a signed-in administrator; the POST also needs the CSRF header.

| Route | Returns |
|---|---|
| `GET /v1/audits` | Every audit's definition, availability for this session, last run and running run |
| `POST /v1/audits/{auditId}/runs` | 202 with the new run, or 200 with the run already in progress |
| `GET /v1/audits/{auditId}/runs` | The last runs, without rows |
| `GET /v1/audit-runs/{runId}` | One run with its rows (no rows while running) |
| `GET /v1/audit-runs/{runId}/export.csv` | CSV of the run; `view`, `search`, `sort`, `direction` and `f.<column>` filter it |
| `GET /v1/settings/audits` | The saved scope and the domains to choose from, with how many accounts each holds (section 9) |
| `POST /v1/settings/audits` | Saves `{ "domains": [...] }` and returns what the GET returns; 400 unless the body is a list of up to 100 domain names |

| `GET /v1/audits/{auditId}/trend?days=365` | The audit's trend points in the last 1 to 3,650 days (default 365), oldest first (section 11) |
| `GET /v1/settings/audit-schedule` | Every audit with `off`, `daily` or `weekly`, and whether collection is on (section 12) |
| `POST /v1/settings/audit-schedule` | Saves `{ "audits": { "<auditId>": "daily" } }` for the audits named and returns what the GET returns; 400 unless every value is `off`, `daily` or `weekly`; 403 when an audit turned on is one the saving administrator could not run |

CSV is RFC 4180 with a byte-order mark. Text that a spreadsheet would treat as a formula
(starting with `=`, `+`, `-`, `@`, a tab or a carriage return) gets a leading apostrophe.

## 7. Adding an audit

An audit is one object implementing `Audit` in `packages/audits/src/definition.ts`, added to
`builtInAudits`:

- `definition`: id (`area.name`), title, description, category, what a row is, the columns
  (`person`, `text`, `enum`, `date`, `days`, `number`, `boolean`, `list`; `facet` makes an enum a
  filter, `hidden` leaves it out of the table by default), the preset views as tag sets, the
  default view and sort, and requirements in plain words.
- `packs` and `optionalPacks`: the Graph read packs it needs and the ones that only add detail.
  The Graph client it receives can reach only those packs' endpoints.
- `scoped`: true when the audit covers only the account domains the administrator chose (section
  9). A scoped audit keeps only the accounts `inAuditScope` accepts before it computes rows,
  tiles and notes, starts its notes with `scopeNote`, and has the shared `domainColumn`.
- `forbidden`: what to tell the administrator when Microsoft returns 403.
- `run(context)`: read, then return rows (each with tags), summary tiles, notes, and whether the
  read was complete. Parse provider records leniently and count what could not be read; never
  fail a run for one odd record.

The page, the CSV export and the API need no change. Tests go in `packages/audits/test`; the
definitions test checks that every audit's views, sort column and tiles agree with its columns.

## 8. Tests

`packages/audits/test`: the inactivity rules, every audit against a synthetic tenant through the
real Graph client and endpoint allowlist, refusals, partial reads and field fallbacks.
`packages/contracts/test/audits.test.ts`: views, filters, search, sorting and CSV. `apps/api/test/audits-api.test.ts`:
sessions, CSRF, availability, one run at a time, history, export, failures, the time limit,
recovery of interrupted runs, shutdown and pruning. `apps/api/test/audits-scope-api.test.ts`: the scope settings, their ledger entry and the scope applied to runs. `apps/web/test/audits.test.ts`: the catalog and the scope section.

Trends and schedules: `packages/audits/test/trends.test.ts` (points written with the result, none
for failed or interrupted runs, kept through pruning, tile ids the same in SQL and TypeScript,
comparability, the saved schedule and when an audit is due), `packages/storage/test/audit-trend-migration.test.ts`
(the backfill on PGlite; the same suite runs on PostgreSQL in `postgres.test.ts`),
`apps/api/test/audit-insights-api.test.ts` (the trend and schedule routes),
`apps/api/test/audit-schedule.test.ts` (the scheduler on a fake clock, and in the app with the real
runner) and `apps/web/test/audit-insights.test.ts` (the trend panel and the schedule control).

## 9. Scope

An administrator can limit the account audits to some of the tenant's domains. The scope is a
list of domain names, up to 100, saved once per tenant under the key `audits.scope` in
`prism365_preference`. An empty list means every domain, and so does a missing or unreadable value.

A domain is the text after the last @ in an account's user principal name, in lower case. The
Inactive accounts, Licenses to reclaim and MFA registration audits use the scope: an account is
covered when its domain is in the list, and an account with no principal name is covered only when
the list is empty. Guest accounts and License capacity ignore the scope. A guest's principal name
ends with the tenant's own domain, so guests appear under that domain in the scoped audits.

An audit reads the scope once, when its run starts, and the run stores it. It filters the accounts
before it counts anything, so the tiles, the table and the CSV export describe the same accounts.
A scope saved while a run is reading Microsoft applies from the next run.

A scoped result always says what it covered and what it left out:

- its notes start with "Limited to accounts in: a.example, b.example." and say how many accounts
  in other domains were left out;
- the MFA audit counts admins without MFA in the left-out domains and shows them in their own
  critical tile, "Admins without MFA outside the scope", and its in-scope admin tile is never shown
  as all clear under a scope; Inactive accounts does the same for accounts with sign-in attempts
  and no success. An emergency-access account often sits on the tenant's onmicrosoft.com domain,
  so these are the accounts a scope must never hide;
- its results page shows "Limited to N domains" for the scope that run used, and says when the
  saved scope has changed since;
- its CSV file name ends in "-limited-to-N-domains".

Each scoped audit also has a Domain column, hidden by default, with a filter above the table. The
filter narrows what is on screen to one domain without changing the saved scope.

The Audit scope section at the top of the Audits page shows the saved scope and a Change button,
which opens a list of domains with the number of accounts in each. Every scoped run stores the
domains it read before its scope was applied, so the list keeps offering domains the current scope
leaves out, and a domain that appears later shows up after the next run. Runs made before this
was stored fall back to the domains in their rows. Nothing is read from Microsoft to build the
list, it shows at most 1,000 domains (the largest), and until a scoped audit has run, the page
asks for a first run of Inactive accounts.

Saving the scope writes an access-ledger entry, `audits.scope.changed`, with the number of domains
and the time, before the scope is saved. It does not name them. The scope needs no Microsoft
permission. The run's scope and domain counts are two columns of `prism365_audit_run`, added in
place by migration 004.

## 10. Identity audits

Both audits in the Identity group read endpoints that Microsoft limits to one request per second
for each tenant, and the same limit covers Conditional Access policies, named locations, risky
users and risk detections together. Microsoft answers a request over the limit with HTTP 429 and
no `Retry-After`, which leaves the Graph client to guess at a wait of half a second, one second
and two seconds. So these audits space their own requests: the first goes straight out and each
later one to these endpoints waits 1.1 seconds. A Conditional Access run that needs named locations
takes a little over a second longer for it, and a risky users run waits once per page of 500. The
wait stops the moment the run's time limit passes. Two of these audits running at the same time
share the limit and can still be answered with a 429; the client retries, and a run that is still
refused ends as throttled. Sources: [throttling limits](https://learn.microsoft.com/en-us/graph/throttling-limits).

### Conditional Access policies

Reads `GET /identity/conditionalAccess/policies` with the `identity.conditional-access.read` pack
(`Policy.Read.All`). The request sets no page size, because Microsoft does not document a maximum,
and follows the continuation link. It reads `GET /identity/conditionalAccess/namedLocations` only
when a policy points at a named location, to show the location by name. With `directory.users.read`
granted it looks up the display name of each excluded user, one request per user and at most 50
users, asking for the name fields only. With `tenant.licenses.read` granted and an empty policy
list, it reads the tenant's subscriptions to say whether any of them includes Microsoft Entra ID P1
(the `AAD_PREMIUM` service plan, which the P2 product also carries). Not scoped: a row is a policy.

One row per policy:

- Policy and State: On, Report only or Off. A state this version does not know is shown as Unknown
  and counted in none of the three tiles.
- Applies to: all users, or counts of users, groups and roles (with names for the common administrator
  roles), guests and external users with their types, or workload identities.
- Leaves out: excluded users, groups and roles as counts, with the users' names when the session
  has granted `User.Read.All`. An account Microsoft no longer has is counted as deleted or unknown.
- Apps: all apps, Office 365, Microsoft admin portals, a count of specific apps, user actions, and
  how many apps are excluded.
- Conditions: client app types, platforms, locations by their named-location name, and sign-in,
  user and service principal risk levels. Conditions that match everything are left out.
- Grant: block, or what must be satisfied, with "one of" when the operator is OR. An authentication
  strength shows by name.
- Session: sign-in frequency, persistent browser, app enforced restrictions, Conditional Access App
  Control and disabled resilience defaults.
- Modified, and Created (hidden), plus counts of excluded users, groups and roles (hidden).

Views are Enabled, Report only, Off, Blocks legacy authentication, Requires MFA and Has exclusions.

Tiles:

- Policies enabled, Policies in report-only and Policies off.
- No enabled policy blocks legacy authentication (critical, shown only when true). A policy counts
  when it is on, applies to all users, lists both Exchange ActiveSync and other clients as client
  app types, and blocks. A report-only policy, a policy for some users and a policy for one of the
  two client types do not count. A block on all client apps does not count either, because it is
  not aimed at legacy clients and another condition usually narrows it.
- No enabled policy requires MFA for administrators (warning, shown only when true). A policy
  counts when it is on, its grant includes MFA or an authentication strength, it covers all users or
  the Global Administrator role (template ID `62e90394-69f5-4237-9190-012177145e10`), and it does
  not exclude that role.
- Users excluded from MFA policies: the distinct users excluded from at least one policy that is
  on and requires MFA. Break-glass (emergency access) accounts are often here on purpose, so the
  tile asks for a check, not a fix. The run's notes say the same, because a tile has no help text.

An empty policy list gets a note that says whether the tenant has Entra ID P1, does not, or could
not be checked. Without proof that Conditional Access is available, the two "No enabled policy"
tiles are left out, because a tenant that cannot have a policy has not missed one.

What it cannot see:

- Report-only policies do not enforce, so every check leaves them out.
- The checks do not weigh apps, locations, risk levels or exclusions. A policy for all users that
  excludes most of them still counts, and Has exclusions and Leaves out are where to look. A grant
  with the operator OR counts as requiring MFA even though another listed control can satisfy it.
- Policies refer to roles by role template ID. Prism365 names sixteen common administrator roles
  and counts the others. It does not say who is in an excluded group or holds an excluded role.
- Security defaults are not Conditional Access policies and are not read. When they are on,
  Microsoft blocks legacy authentication and asks administrators for MFA, so a missing policy may not
  leave the tenant open.
- Nothing is evaluated: there is no what-if, and no sign-in is matched to a policy.
- Microsoft does not document the page size, whether a tenant without P1 gets an error or an empty
  list, or how a list of guest types is returned (a comma-separated string or an array). The audit
  reads both forms of the guest types and treats an empty list as above.

A policy Prism365 cannot read is left out, counted in a note, and makes the run partial. If the
named locations cannot be read the policies are listed with locations unnamed, and the run is
partial.

### Risky users

Reads `GET /identityProtection/riskyUsers?$top=500` with the `identity.risk.read` pack
(`IdentityRiskyUser.Read.All`), a page of 500 at a time (Microsoft's documented maximum) with a
wait between pages, up to 10,000 users. Microsoft documents that this API needs Microsoft Entra ID
P2; what a P1 tenant receives is not documented, so both a refusal and an empty list say "Risky user
data in Microsoft Graph needs Microsoft Entra ID P2." Scoped by the account's domain, with the
Domain column and the scope notes of section 9.

One row per risky user: the user (linked to the person's page; the row's ID is the user's object
ID), risk level (High, Medium, Low, Hidden or None), risk state (At risk, Confirmed compromised,
Remediated, Dismissed, Confirmed safe or None), risk detail as words, when Microsoft last updated
the risk and how many days ago, and whether Microsoft is still processing the user. Users Microsoft
lists as deleted are left out and counted in a note.

Views are At risk (at risk or confirmed compromised), High risk (high level and still at risk),
Remediated or dismissed, and All.

Tiles:

- Confirmed compromised (critical when above zero).
- High risk, still at risk (critical when above zero).
- Medium risk, still at risk (warning when above zero).
- Remediated or dismissed in the last 30 days, by the date Microsoft last updated the risk.
- High risk or compromised outside the scope (critical, shown only when above zero). A scope
  never hides these silently, for the reason the MFA audit gives for administrators without MFA.

What it cannot see:

- Microsoft keeps a risky user until the risk is remediated, so an old entry stays listed until
  someone closes it.
- Hidden means the account is not covered by ID Protection, so Microsoft reports no risk for it.
- The results are Microsoft's detections. Prism365 adds none, reads no risk detections or risky
  sign-ins, and does not say why Microsoft rated a user.
- The risk detail values Microsoft sends only on request are not asked for. If a tenant sends one
  anyway it is shown in words.

## 11. Trends

Every successful run, complete or partial, leaves a trend point: its summary tiles (inactive
accounts, admins without MFA, unused licenses and the rest) with their values, the audit's version,
the scope the run used, whether it was complete, how it was started and when it finished. Points
live in `prism365_audit_trend` (core migration `005-audit-trends.sql`), a tenant table with forced
row-level security. They are written in the same transaction as the run's result, so a result that
could not be stored leaves no point, and a failed or interrupted run never has one. Pruning runs
never touches points and nothing else deletes them: a point is a few hundred bytes, so a daily
audit adds well under a megabyte a year. A point holds counts and the labels the audit gave its
tiles, never an account.

Tiles have no id of their own, so a point keys each one by its label: "Inactive 30+ days" becomes
`inactive-30-days` (ASCII letters and digits kept, anything else one hyphen, at most 60 characters,
"tile" when nothing is left; `auditTileId` in `packages/contracts/src/insights.ts`). A label that
gives the same id as an earlier tile of the same run is dropped from the point.

When migration 005 first runs, the successful runs already stored become the first points, so an
installation that had runs before trends existed starts with a trend. The backfill derives tile ids
in SQL exactly as `auditTileId` does and skips runs that already have a point, so it changes
nothing when the migration runs again on the next start. The migrator connects as a superuser; a
migration role that only owns the tables sees no rows under the forced tenant policy, and then the
backfill adds nothing and points start with the next run.

An audit's page shows a Trend panel once the audit has two points in the last year. It draws one
tile at a time, the first critical or warning tile by default, as a line chart with a title and
description for screen readers, and Show as table lists the points instead. A point is drawn
hollow, with a note saying why, when it does not compare directly with the latest run: it was made
by another version of the audit, it covered another audit scope, or it was a partial result. Every
point is compared with the newest one, because that is the result the page shows. A run that had
no value for the chosen tile leaves a gap in the line.

Reading a trend is not an access-ledger event: points are counts, like the run history without its
rows, and neither reads Microsoft. The route returns at most 1,000 points and says when the range
held more.

## 12. Schedules

An audit can run on its own, daily or weekly. Run automatically on the audit's page has Off, Daily
and Weekly and saves at once; the catalog tags a scheduled audit Daily or Weekly, and the run
history marks scheduled runs.

You can schedule an audit only if you could run it yourself: turning an audit on (or moving it to
another cadence) needs every read pack it requires enabled in `ENABLED_READ_FEATURES` and granted
by your own sign-in, by the Run button's rule, and is refused with the Run button's message
otherwise. Turning a schedule off is always allowed. Scheduled runs then read with the collecting
administrator's sign-in, not yours.

The schedule is one setting per tenant under the key `audits.schedule` in `prism365_preference`:
`{ "audits": { "<auditId>": "off" | "daily" | "weekly" }, "scheduledBy": { "<auditId>": "<actor id>" } }`,
where `scheduledBy` is the administrator who last set each scheduled audit to its cadence. A save
changes only the audits it names, so two pages changing different audits never undo each other.
An id that is not a known audit is dropped from a save and ignored when running, but an entry a
newer build saved for an audit this build does not know is kept when this build saves. A value
that is not a frequency reads as off, so a damaged setting can only stop audits from running.

Scheduled audits run only while collection is on, with the sign-in collection uses: the
administrator who started collecting. The API checks every 5 minutes, and once 30 seconds after
collection starts. An audit is due when its newest run of any outcome, manual or scheduled,
started more than a day (daily) or a week (weekly) ago, or when it has never run. Counting from the
last start means a manual run counts as this period's run, and a scheduled run that failed waits
for its next turn instead of retrying every few minutes. A start in the future, after the clock was
turned back, is not due until a full period after it; the audit's page then says "Next run waits
for <date>: the last run is dated in the future."

Due audits run one at a time, in the catalog's order, through the same runner as a manual run.
Before each one the scheduler reads the schedule and the last start again, so a schedule turned off
or a manual run started meanwhile is respected, and it checks the collecting sign-in again just
before starting, so nothing starts after collection stops. An audit is skipped when any of its
read packs is not enabled or not granted by the collecting sign-in, by the same rule as the Run
button; the server log names the audit and the reason once, and again only when the reason
changes, and the audit stays due. A run reads its own packs and the optional packs the operator
enabled, with the permissions the collecting sign-in granted for them (`auditReadScope`), exactly
as a manual run does.

A scheduled run reads only while its sign-in is still the one collecting. Every token it asks for
is refused once collection has stopped or passed to another administrator, and the run in
progress is stopped the moment that happens, its request in flight included. It closes as failed
with "Collection stopped while this scheduled run was reading." Manual runs are never stopped this
way. A sign-in that stops working while a scheduled run reads (signed out, expired, or refused by
Microsoft) fails that run as a session failure, the way a manual run fails, and ends the round: the
remaining due audits are neither started nor recorded until the next check.

A scheduled run is stored with trigger `schedule` and the collecting administrator as the one who
started it. Its access-ledger entry is the manual run's, `audit.run.started` with the run and the
permissions it reads, plus `"trigger": "schedule"` and `"scheduledBy"`, the administrator who
scheduled it. The entry is written under the collecting administrator, whose sign-in the run reads
with, before anything is read. Scheduled runs raise no notifications: the results are on the
audit's page.

Saving the schedule writes `audits.schedule.changed` before the schedule is saved, with the audits
the save changed and their new cadence (`"changes": [{ "auditId", "frequency" }]`) and how many
audits run on their own after it. It holds audit ids and cadences only.
