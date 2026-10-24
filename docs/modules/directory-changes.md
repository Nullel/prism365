# Module: directory.changes

Status: built (TL-030), first and second adversarial reviews fixed, third review pending. Package: `packages/module-directory-changes`.

## 1. What it answers

Who changed what in the directory, and when. Microsoft Entra writes a directory audit event
for every change to a user, group, role, application, policy or domain: who made it, what it
was made to, and the values before and after. Microsoft keeps these events 7 days without Entra
ID P1 or P2 and 30 days with it. Prism365 keeps every one it collects for 10 years by default
(`DIRECTORY_CHANGES_RETENTION_DAYS`, 365 to 3650), so after a month it is the only place that
still has the history.

The administrator reads it three ways:

- the Directory changes page: a feed grouped by day, filterable by range, category, result
  and text, with the changes the rules flagged on top;
- a Directory changes panel on each person's page: what they changed and what was changed
  on their account;
- findings for the changes that matter for security (section 5), in the same lifecycle as
  every other finding.

## 2. Setup

| Read pack | Graph permission (delegated, admin consent) | Signed-in role |
|---|---|---|
| `audit.directory.read` | `AuditLog.Read.All` | Reports Reader, Security Reader or Security Administrator |

1. Add `AuditLog.Read.All` to the app registration and grant admin consent (the sign-in audit
   uses the same permission).
2. Add `audit.directory.read` to `ENABLED_READ_FEATURES`. The scheduler runs the module only
   when this pack is enabled, even if the permission is already consented for sign-ins.
3. Sign in again, or use the grant link on the page, and start collecting.

Microsoft's pages disagree on licensing: the licensing table lists audit logs for every edition,
while the activity log pages say reading them through Graph needs a premium license. A 403 is
shown as "needs a role such as Reports Reader, and possibly Entra ID P1 or P2".

## 3. Source

`GET /v1.0/auditLogs/directoryAudits`, checked against Microsoft Learn on 2026-10-22
(`directoryaudit-list`, the `directoryAudit` and `appIdentity` resources). Windows are read with
`$filter=activityDateTime ge X and activityDateTime le Y`: Learn documents `eq`, `ge` and `le`
for `activityDateTime`, not `lt`, so neighbouring windows share their boundary millisecond and
records in it are read twice and stored once. Nothing depends on the order Microsoft returns
records in. `$top` is documented without a maximum; Prism365 asks for 999 and drops to 100 for
the rest of a run if Microsoft refuses the request or a page is larger than the client accepts.

Each record is checked on its own. One without an id or a readable time is left out and counted
(the run shows as partial with the number left out); everything else is tolerated, and unknown
fields are dropped. What is stored is bounded: names cut to 256 characters, property values to
400 after unquoting Microsoft's JSON-quoted values (with a flag saying a value was cut), at most
50 targets with 50 changed properties each, and 20 additional details. Provider text is cleaned
and never rendered as HTML. The actor's IP address is read and not kept, and additional details
that describe the client (User-Agent, client IP address, browser or device information) are
left out.

## 4. Collection

Collector `audits`, every 10 minutes, stopped by the scheduler at 120 seconds. Each run reads for
at most 25 seconds and starts no request after that, so the last request has its 90 seconds to
finish. Each request the module makes starts at least 2.2 seconds after the previous one
finished, across runs, so the module's requests stay within the 5 per 10 seconds Microsoft allows
per app and tenant. When Microsoft answers 429, the Graph client retries inside that one request
after the wait Microsoft gives; those retries are Microsoft's pace, not the
module's. A wait carried over from an earlier run is never longer than one interval, so a clock
that stepped back cannot hold collection.

- The first run starts 30 days back (Microsoft's longest retention) and reads forward in windows
  of up to a day, asking the scheduler to run again within seconds while history remains. If
  every read fails before anything was read (a missing role, say), history simply starts later;
  no finding is raised for it.
- Two lanes read the log. The main lane reads new records from where it stopped. The settle
  lane reads every window again once it is the late-arrival margin old: 240 minutes by default,
  60 to 1440 with `DIRECTORY_CHANGES_LATE_MINUTES`. That includes the whole history after the
  first load, so one empty answer from Microsoft, which Learn says can happen under load, is
  never the only read of a window. Each lane saves its own position, so a run that ends early
  never makes the next start over. While the settle lane is behind, the main lane leaves it a
  third of each run. With time left, a run also re-reads, newest first, what earlier runs read
  that has not settled yet, so most late records appear within a run; this best-effort pass
  saves nothing.
- A lane moves only past a window it read completely, or one recorded as a gap, and saves its
  place after each one. A run that stops at its time budget keeps the position at the end of the
  last complete window. Before reading a window the lane records that it is reading it; a run
  stopped mid-window is seen by the next one, which counts that window as failed.
- A window holding more than 5,000 records is split in half and read again from the same place,
  down to a single millisecond (Microsoft's filter is millisecond-precise). A window that fails
  is halved and read again at once, in the same run while it has time: one that times out, that
  Microsoft answers with a server error or a request to wait longer than the client waits (a
  deferred 429), that Microsoft refuses (400) or answers too large at both page sizes, whose
  answer is malformed or pages in a loop, or that could not be finished in a run that gave it at
  least half its time. A run still narrowing one down asks to run again within seconds. Only a
  window that cannot be split further is given up: one that still holds too much keeps what it
  read and is recorded as a gap; one that keeps failing is tried three times, the last at least
  10 minutes after the first (on a later scheduled run), then recorded as a gap that says why,
  and the lane moves past it. The wait keeps a burst of failures that has nothing to do with the
  window from using up its tries within seconds. The window after it, which starts on its last
  millisecond (both bounds are inclusive), counts from the same first failure. Windows lie on a grid of their own size, and
  gap findings are keyed by the minute they start in, so the same span is one finding however
  often it is read.
- A failure is blamed on the window only while Microsoft answers other requests: to halve a
  window, anything answered in the run; to count a try against a millisecond, something
  answered since the last failure. Otherwise a one-millisecond read of the present tells. When
  that fails too, Microsoft is failing rather than the window (an outage, or throttling): the run
  fails with that error and changes nothing, so an outage leaves no gap and the next try waits
  for the cadence. A 401, a 403 or a missing credential also fails the run and keeps the lane
  where it is.
- A gap that the settle lane later reads whole, when that read returns changes inside it, is
  retitled "A gap in directory changes was read later" and says nothing is missing; a person
  resolves it. A read that comes back with nothing leaves the gap open, since Microsoft can answer
  empty under load. The best-effort re-read stores what it finds in a gap but does not retitle
  it.
- Events are stored by id; a record read again changes nothing, and neither do its subjects.
- If the main lane falls behind what Microsoft still keeps after it had read something,
  because collection was stopped or every read failed for that long, the span is recorded as
  one gap finding that widens while the lane stays behind.
- If the settle lane falls behind what Microsoft still keeps, it moves to what Microsoft keeps.
  The part of the span it skipped that the main lane read while it was fresh (younger than the
  margin) is recorded as "Late directory changes may be missing": a change Microsoft published
  late there may never have been read. What the main lane read when it was already old (the
  history load, a catch-up) counts as settled. For both kinds, a new span that starts within the
  margin of the last one's end widens the same finding.
- A saved position in the future (a clock that ran ahead and came back) brings the main lane
  back to now. When the clock ran ahead by more than the margin, the settle lane reads
  everything Microsoft still keeps again, since changes made meanwhile may lie behind it.
- A cursor that cannot be read (garbage, or from a later version) starts the history again;
  repeats change nothing.

## 5. Rules

Each rule has an id and a version, judges one newly stored event (an event stored by a run that
stopped before its rules ran is judged by the next run), and produces at most one finding,
keyed by the event id unless the table says otherwise. Only successful changes are judged: a
failed attempt changed nothing and stays in the feed. Activity and category names are compared
without case, and Microsoft's en dashes as hyphens. Names are from Microsoft's "Microsoft Entra
audit log activity reference"; role template IDs from "Microsoft Entra built-in roles".

| Rule | Category | Activities | Severity | Subject |
|---|---|---|---|---|
| `role.privileged-assigned` | RoleManagement | Add member to role; Add eligible member to role; Add scoped member to role; Add member to role scoped over Restricted Management Administrative Unit; Add member to role in PIM completed (permanent); Add member to role in PIM completed (timebound); Add member to role completed (PIM activation) | critical for Global Administrator; high for Privileged Role, Privileged Authentication, Security, Conditional Access, Exchange, SharePoint, User, Application, Cloud Application, Authentication and Intune Administrator | the person (or group or app) added |
| `app.credential-added` | ApplicationManagement | Add service principal credentials; Update application - Certificates and secrets management (when a key ID appears that was not there before, or the key list is missing or too long to compare) | high | the application |
| `app.consent-granted` | ApplicationManagement | Consent to application | high when ConsentContext.OnBehalfOfAll is true (or, without it, IsAdminConsent), medium otherwise | the application |
| `policy.conditional-access-changed` | Policy | Add, Update, Delete Conditional Access policy | medium; high for a deletion | the policy |
| `domain.federation-changed` | DirectoryManagement | Set domain authentication; Set federation settings on domain; Update domain when a changed property is named as an authentication or federation setting | critical | the tenant |
| `auth-method.removed-by-admin` | UserManagement | Admin deleted security info; Update user that empties StrongAuthenticationMethod (an administrator making the person register MFA again). Not when the actor changed their own account | medium | the person |
| `collection.gap` | | | medium (operations) | the collector |

Only PIM's own "Add member to role completed (PIM activation)" record makes a role activation.
The "Add member to role" record the directory writes for it, made by the service principal named
MS-PIM (a name seen in tenants rather than documented), counts in the same finding when it
shares that record's correlation id. Activations are one finding per account, role and day,
counting each one. Everything else is one finding per event, titled as an assignment: permanent
and timebound assignments made in PIM, eligible assignments, and an MS-PIM record with no
activation beside it. So an assignment never folds into an activation someone has already
acknowledged. When the directory's record is judged before PIM's own record is stored (Microsoft
published it later, or it sits in the next window), it stands as its own assignment finding.
Learn's reference has no federation-specific "Update domain" activity, so that match depends on
the property name.

Evidence is the event id, time, activity, category, result, actor and target, plus the role,
consent scope or domain where a rule uses one. Each finding says what the rule cannot see:
roles granted through membership of a role-assignable group appear only as a group change (the
rules do not follow group membership to roles), the record holds only 400 characters of a
policy, and Microsoft does not always name the method removed. A directory name that trips the
finding language policy is left out of the finding rather than failing the run.

## 6. Routes

- `GET /v1/changes?days=&category=&result=&search=&before=&limit=`: newest first. `days` 1 to
  3650 (default 7), `result` `success` or `failure` (anything that did not succeed), `search` up
  to 200 characters over the activity, the actor's name and sign-in name, and the targets' names
  and sign-in names (not the changed values), `limit` 1 to 200 (default 100), `before` the
  previous page's `next`. The search and the stored text are folded the same way (compatibility
  decomposition, accents and other marks removed, lower case, final sigma as sigma), so "ilker"
  finds "İlker" and "ΟΔΥΣ" finds "ΟΔΥΣΣΕΥΣ". Returns `items`, `next`, counts by category for the range and coverage
  (`from`, `through`, `lastCollected`, `phase`).
- `GET /v1/people/:userId/changes?before=&limit=`: the person's changes as actor or target,
  listed once each, with how they were involved. `limit` 1 to 200 (default 20).
- Findings come back from `GET /v1/findings?module=directory.changes`.

Both routes check the session (401), answer 409 in plain words when module storage or this
module is unavailable, validate every input (400), and write an access-ledger entry before
reading: `changes.viewed` and `person.changes.viewed`, each with `datasets: ["directoryAudit"]`
and counts only (the range in days), never a name, a search text or a value.

## 7. Storage and retention

`mod_directory_changes_event` (one row per event, targets and details as bounded JSON, a
folded search text of up to 8,000 characters, a `processed` flag for the rules) and `mod_directory_changes_subject`
(each account or object an event involves, as actor or target, indexed by subject and time, so a
person's changes are a plain lookup). Both are tenant-scoped. The scheduler's daily retention
deletes rows older than `DIRECTORY_CHANGES_RETENTION_DAYS` (default 3650, 10 years). What each
row holds is listed in `docs/PRIVACY.md`.

## 8. Limits

- Records Microsoft publishes later than the late-arrival margin (240 minutes by default) are
  not picked up.
- Learn notes that the API can answer with empty results under heavy load. Every window the main
  lane read is read again by the settle lane, so a record is lost only if both reads come back
  empty. A span the main lane gave up is read only by the settle lane (and, without saving
  anything, the best-effort re-read); if that read comes back empty, its gap finding stays open.
- A millisecond Microsoft never answers for is given up and its changes are missing; everything
  around it is read.
- A record exactly on a millisecond boundary sits in the windows on both sides of it, so a gap
  around it can span two milliseconds.
- Directory audit events say what Microsoft recorded about a change. They are not a measure of
  anyone's work, and the feed never ranks people.

## 9. Tests

`packages/module-directory-changes/test`: history across several runs with `more`; request
pacing on every lane, the page-size fallback and back-to-back runs, and a clock that stepped
back; late records on the best-effort re-read and on the settle lane with no time to spare; the
240-minute default; repeats; a budget running out mid-window; splitting below a minute; a
one-millisecond gap found once; a window that never answers and one that keeps failing, both
narrowed at once and given up on a scheduled run ten minutes on, the second healed when the
settle lane reads its change; a burst of failures shorter than ten minutes giving nothing up; gap findings left open on an empty read and not healed by the best-effort re-read;
answers too large, refused or paging in a loop at both page sizes; gateway errors on wide windows
caught up in one poll; an outage that changes nothing; throttling that starts mid-run; empty
answers in the history load and a catch-up; a run stopped mid-request; strikes cleared; refused
access raising nothing; widening findings for a span Microsoft dropped, for the span the settle
lane skipped, and at the retention edge; a clock that ran ahead; the settle lane's third of the
run; an acknowledged PIM activation followed by a permanent assignment; unreadable records;
garbage cursors; the smaller page size; rules restarting; the language-policy fallback; settings
and retention. Each rule against its near misses, including PIM activations and assignments,
cleared MFA methods (with a list cut for length) and domain updates; source normalization,
client details and search text; feed and person queries, folded search and list bounds.
`apps/api/test/changes-api.test.ts`, `apps/web/test/changes.test.ts`, and
`scripts/demo/graph/directory-audits.test.ts` for the synthetic log the demo is recorded from.
