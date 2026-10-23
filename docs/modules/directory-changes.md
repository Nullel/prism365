# Module: directory.changes

Status: built (TL-030), adversarial review pending. Package: `packages/module-directory-changes`.

## 1. What it answers

Who changed what in the directory, and when. Microsoft Entra writes a directory audit event
for every change to a user, group, role, application, policy or domain: who made it, what it
was made to, and the values before and after. Microsoft keeps these events 7 days without Entra
ID P1 or P2 and 30 days with it. Prism365 keeps every one it collects for as long as it runs, so
after a month it is the only place that still has the history.

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
and never rendered as HTML. The actor's IP address is read and not kept.

## 4. Collection

Collector `audits`, every 10 minutes, stopped by the scheduler at 120 seconds. Each run reads for
at most 25 seconds and starts no request after that, so the last request has its 90 seconds to
finish; requests start at least 2 seconds apart, inside Microsoft's limit of 5 directory audit
requests per 10 seconds per app and tenant (429 answers are retried by the Graph client).

- The first run starts 30 days back (Microsoft's longest retention) and reads
  forward in windows of up to a day, asking the scheduler to run again within seconds while
  history remains.
- Two lanes read the log. The main lane reads new records from where it stopped. The settle lane follows
  it at the late-arrival margin (120 minutes): every window is read again once it is that old, so
  a record Microsoft publishes up to two hours late is collected. A main-lane window read when it
  was already that old counts as settled. Each lane saves its own position, so a run that ends
  early never makes the next start over. With time left, a run also re-reads, newest first, what
  earlier runs read that has not settled yet, so most late records appear within a run; this
  best-effort pass saves nothing.
- A lane moves only past a window it read completely, and saves its
  place after each one. A run that stops at its time budget keeps the position at the end of the
  last complete window. Before reading a window the lane records that it is reading it; a run
  stopped mid-window is seen by the next one, which counts that window as failed.
- A window holding more than 5,000 records is split in half and read
  again from the same place, down to one minute. A window that times out, or that could not be
  finished in a run that gave it at least half its time, is halved for the next run. A
  one-minute window that still holds too much keeps what it read and is recorded as a gap; one
  that fails on three runs in a row is recorded as a gap and passed. Windows lie on a grid of
  their own size, and gap findings are keyed by the minute they start in, so the same busy
  minute is one finding however often it is read.
- Events are stored by id; a record read again changes nothing, and neither do its
  subjects.
- If collection was stopped for longer than Microsoft keeps the log, the span
  that fell out of it is recorded as a gap.
- A cursor that cannot be read (garbage, or from a later version) starts
  the history again; repeats change nothing.

## 5. Rules

Each rule has an id and a version, judges one newly stored event (an event stored by a run that
stopped before its rules ran is judged by the next run), and produces at most one finding,
keyed by the event id. Only successful changes are judged: a failed attempt changed nothing and
stays in the feed. Activity and category names are compared without case, and Microsoft's en
dashes as hyphens. Names are from Microsoft's "Microsoft Entra audit log activity reference";
role template IDs from "Microsoft Entra built-in roles".

| Rule | Category | Activities | Severity | Subject |
|---|---|---|---|---|
| `role.privileged-assigned` | RoleManagement | Add member to role; Add eligible member to role; Add scoped member to role; Add member to role scoped over Restricted Management Administrative Unit | critical for Global Administrator; high for Privileged Role, Privileged Authentication, Security, Conditional Access, Exchange, SharePoint, User, Application, Cloud Application, Authentication and Intune Administrator | the person (or group or app) added |
| `app.credential-added` | ApplicationManagement | Add service principal credentials; Update application - Certificates and secrets management (when a key ID appears that was not there before, or the values are too long to compare) | high | the application |
| `app.consent-granted` | ApplicationManagement | Consent to application | high when ConsentContext.OnBehalfOfAll is true (or, without it, IsAdminConsent), medium otherwise | the application |
| `policy.conditional-access-changed` | Policy | Add, Update, Delete Conditional Access policy | medium; high for a deletion | the policy |
| `domain.federation-changed` | DirectoryManagement | Set domain authentication; Set federation settings on domain | critical | the tenant |
| `auth-method.removed-by-admin` | UserManagement | Admin deleted security info (not when the actor removed their own) | medium | the person |
| `collection.gap` | | | medium (operations) | the collector |

Evidence is the event id, time, activity, category, result, actor and target, plus the role,
consent scope or domain where a rule uses one. Each finding says what the rule cannot see: PIM
activations also appear as role additions, roles granted through groups appear as group changes,
the record holds only 400 characters of a policy, and Microsoft does not always name the method
removed. A directory name that trips the finding language policy is left out of the finding
rather than failing the run.

## 6. Routes

- `GET /v1/changes?days=&category=&result=&search=&before=&limit=`: newest first. `days` 1 to
  3650 (default 7), `result` `success` or `failure` (anything that did not succeed), `search` up
  to 200 characters over the activity, the actor and the targets' text, `limit` 1 to 200
  (default 100), `before` the previous page's `next`. Returns `items`, `next`, counts by
  category for the range and coverage (`from`, `through`, `lastCollected`, `phase`).
- `GET /v1/people/:userId/changes?before=&limit=`: the person's changes as actor or target,
  listed once each, with how they were involved. `limit` 1 to 200 (default 20).
- Findings come back from `GET /v1/findings?module=directory.changes`.

Both routes check the session (401), answer 409 in plain words when module storage or this
module is unavailable, validate every input (400), and write an access-ledger entry before
reading: `changes.viewed` and `person.changes.viewed`, each with `datasets: ["directoryAudit"]`
and counts only (the range in days), never a name, a search text or a value.

## 7. Storage and retention

`mod_directory_changes_event` (one row per event, targets and details as bounded JSON, a
`processed` flag for the rules) and `mod_directory_changes_subject` (each account or object an
event involves, as actor or target, indexed by subject and time, so a person's changes are a
plain lookup). Both are tenant-scoped. There is no retention rule: the history is kept until the
operator deletes it (see `docs/PRIVACY.md`).

## 8. Limits

- Records Microsoft publishes more than two hours late are missed.
- Learn notes that the API can answer with empty results under heavy load; a window read that
  way during the history load is not read again unless it is within the margin.
- Directory audit events say what Microsoft recorded about a change. They are not a measure of
  anyone's work, and the feed never ranks people.

## 9. Tests

`packages/module-directory-changes/test`: history across several runs with `more`, request
pacing, late records on the best-effort re-read and on the settle lane with no time to spare,
repeats, a budget running out mid-window, splitting, a one-minute gap found once, a window that
never answers, unreadable records, garbage cursors, access failures, the smaller page size,
stopped runs, rules restarting, collection off too long, and the language-policy fallback; each
rule against its near misses; source normalization and bounds; feed and person queries.
`apps/api/test/changes-api.test.ts`, `apps/web/test/changes.test.ts`, and
`scripts/demo/graph/directory-audits.test.ts` for the synthetic log the demo is recorded from.
