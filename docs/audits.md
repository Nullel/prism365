# Audits

Status: built (TL-025). Packages: `packages/audits` (the audits and their run store),
`apps/api/src/audits.ts` (the runner), `apps/web/src/Audits.tsx` (the page). Decision: ADR-0008.

An audit is a check an administrator runs when they need it. It reads Microsoft once, with the
signed-in administrator's delegated session, and keeps the result so it can be reopened, compared
with earlier runs and exported. Nothing runs on a schedule, and no audit changes the tenant.

## 1. The built-in audits

| Audit | Question it answers | Needs |
|---|---|---|
| Inactive accounts (`accounts.inactive`) | Which accounts have not signed in, and for how long? Views for 7, 30, 60, 90, 120, 180 and 365+ days, never signed in, and attempts with no success. | `User.Read.All`, `AuditLog.Read.All`, Microsoft Entra ID P1 or P2, a role such as Reports Reader, Security Reader or Global Reader |
| Licenses to reclaim (`licenses.reclaim`) | Which licensed accounts are disabled or 30+ days inactive, and was each license assigned directly or by a group? | `User.Read.All`; sign-in dates as above, or it lists disabled accounts only |
| MFA registration (`mfa.registration`) | Who has no strong method registered, which admins lack MFA, who relies on text or calls? | `AuditLog.Read.All`, a role such as Reports Reader, Security Reader, Security Administrator or Global Reader |
| Guest accounts (`guests.review`) | Which invitations were never accepted, which guests never signed in or went quiet, from which domains? | `User.Read.All`; sign-in dates as above, or invitation states only |
| License capacity (`licenses.capacity`) | Which products have no seats left, more assigned than owned, an expired or suspended subscription, or unassigned seats? | `LicenseAssignment.Read.All`, a role such as Directory Readers, Global Reader or License Administrator |

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
installation that enabled `audit.signins.read` needs nothing more. A pack's permission must be
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
