# This week

Status: built (TL-036), review pending. Route: `GET /v1/digest`. Code: the sections in
`apps/api/src/digest.ts`, the route in `apps/api/src/routes/digest.ts`, the contract in
`packages/contracts/src/digest.ts`, the page in `apps/web/src/Week.tsx`.

## 1. What it answers

What changed since you last looked. The page gathers six things for the last day, 7, 14 or 30
days, each item linking to where it can be looked at:

- new sign-in findings, and the most severe ones still open;
- directory changes, and the ones the rules flagged;
- accounts that went quiet, from the Inactive accounts audit;
- audit tiles that moved;
- Microsoft deadlines and service incidents;
- how collection and the scheduled audits did.

It reads only what Prism365 already stored. Opening the page asks Microsoft nothing, needs no
Microsoft permission and adds no table. Its figures are as fresh as collection and the last audit
runs, which is why the page says "Built from what Prism365 collected" on top.

## 2. Using the page

This week is in the sidebar, right after Overview. The range buttons choose Last day, Last 7
days, Last 14 days or Last 30 days (the page opens on 7). A range runs from its start, included, to
the moment the page is opened, not included. A panel with nothing to report says so in one line
("No new sign-in findings this week."). A panel that could not be read says why, in a different
colour, and is never worded as empty: nothing happened and nothing was looked at are different
things.

Names and titles open the page that owns them. A finding about a person opens that person's page,
and any other finding opens the Sign-in audit or Directory changes page. An account opens its
page. An audit opens the Audits page on that audit. Message center posts and incidents open
Notifications. The Inactive accounts audit opens on its own default view, 30+ days, which is the
view the quiet accounts come from.

## 3. What each section reads

### Sign-in findings

The findings of the sign-in audit module. Opened is findings whose first-seen time is in the range,
counted by severity. Still open is the ten most severe findings in the open state, whenever they
were opened (most severe first, then by score, then newest), with how many there are in all.
Acknowledged findings are not counted as open, as on the Overview. Resolved is findings resolved in
the range, by when they were resolved. The module's own collection gap findings are left out and
counted under Collection.

### Directory changes

The stored directory audit events in the range, by category (the first 50), and their total.
Flagged is the directory change module's findings opened in the range, most severe first, twenty
listed with the total, without its collection gaps. The panel ends with the change feed's coverage:
the day history reaches back to and when it was last collected. The note says when collection
started partway through the range, or when the first load of history is still running.

### Accounts that went quiet

Compares two stored results of the Inactive accounts audit: the newest that finished, and the newest
that started at or before the start of the range. An account is listed when it is in the 30+ days
view of the newer result and was not in that view of the older one. An account the older result did
not list at all is listed too, since the audit cannot say it was active. Fifty are listed, longest
quiet first, with the total. Each item has the person's id, name, sign-in name and days quiet.

The comparison needs an older result that is still kept. Prism365 keeps the newest 10 successful
results of each audit, so a daily schedule reaches back about 10 days and a weekly one about 10
weeks (a result started by hand counts like any other). Without one, the panel says so and suggests
a schedule. It also says when the newest result is itself older than the range.

The two results are compared even when they do not compare directly. The panel says so when they
come from different versions of the audit, covered different domains (the audit scope) or one was
partial, because accounts can then be listed only because of that.

### Audit changes

Every audit keeps one trend point per successful run, and points are never deleted, so this section
reaches back as far as the audit's first run. For each audit it takes the newest point at or before
the start of the range and the newest point before the end, and lists the audits with a tile whose
value differs: the value then, the value now and the change. Critical and warning tiles come first,
and audits with such a tile before the others. A tile that only one of the two results has is not
compared. Each audit says when its two results do not compare directly, for the reasons above, and
the note counts the audits that could not be compared at all (no point from before the range, or
none newer).

### Microsoft

Message center posts that ask for action: Prism365's notification items from the Message center
module that are still active, not marked done and not muted, with an act-by date from now up to,
not including, 30 days ahead, soonest first (twenty listed, with the total). Service incidents:
those first seen in the range, those no longer active whose last change was in the range (resolved
by Microsoft, or no longer listed by it), and the ones open now, most severe first (twenty listed).
The two parts are available on their own, so a missing Service health pack leaves the posts in
place.

### Collection

Whether collection is on, each collector's last successful run and newest run (Prism365 keeps 30
days of run history, so an older success shows as none), the collector runs that failed in the range
with the categories of the five newest failures, the gaps the collectors recorded in the range (the
sign-in and directory change gap findings, any state), and the audits scheduled daily or weekly with
when each last started.

## 4. The route

`GET /v1/digest?days=7`. `days` is 1, 7, 14 or 30 and defaults to 7; anything else is a 400 with a
plain message. The route needs a signed-in administrator (401 otherwise) and reads in one
transaction. It answers 200 even when part of the installation is off:

- a module that is not installed, or whose read pack is not enabled in `ENABLED_READ_FEATURES`, makes
  its section `unavailable` with the reason (a pack that is off collects nothing new, so what is
  stored would be stale);
- a module that is on but has never collected, and has stored nothing, is `unavailable` too;
- audit results, run history and findings are core data, so those sections stay available whenever
  module storage exists;
- without module storage every section is `unavailable` and nothing is read.

Each section has a `state` (`ok`, `empty` or `unavailable`) and a one-line `note`.

Before the first read the route writes an access-ledger entry, `digest.viewed`, with the range in
days and the names of the sections it reads (`signInFindings`, `directoryChanges`, `quietAccounts`,
`auditDrift`, `microsoft`, `collection`; a section whose source is off is not read, so not named),
never a person, a title or a value. If the ledger cannot record it, nothing is read. When there is
no module storage there is nothing to read, so there is no entry.

## 5. Limits

- Everything is stored data. A change Microsoft made but Prism365 has not collected yet is not here,
  and the Collection panel shows how long ago each collector last succeeded.
- "Opened" uses a finding's first-seen time and an incident's first-seen time, which is when
  Prism365 first recorded it.
- The older Inactive accounts result decides what "went quiet" means: the range is only roughly the
  range chosen, and the panel names the date of the result it compared with.
- Counts of failed runs cover the 30 days of run history, which is also the longest range.
- It never ranks people or measures anyone's work. Findings, changes and quiet accounts are
  evidence to look at, as everywhere else in Prism365.
- In the demo the page is replayed from the recording for the four ranges, with its dates moved
  forward. It does not follow what the demo's visitor changes.

## 6. Tests

`apps/api/test/digest.test.ts`: every section on a real (PGlite) database with seeded rows. The
edges of a range (a row exactly at the start counts, one exactly at the end does not), severity
counts, open findings and their order, resolved by resolution time, flagged changes without gaps,
coverage notes, the quiet accounts diff (went quiet, left the view, never in the view, not listed
before, the baseline at and just after the start, a future result, more than fifty, damaged rows,
every comparability flag), audit changes (the points used, changed tiles only, tone order,
comparability, tiles in one point only, audits it cannot compare), the act-by window, incidents,
partial availability, failed runs and categories, gaps and scheduled audits.
`apps/api/test/digest-api.test.ts`: the session, the range, the ledger entry before the first read,
a failing ledger, a missing module, a pack that is off, and no storage.
`apps/web/test/week.test.ts`: every section with data, empty and unavailable, where each item opens,
and the replay. `scripts/demo/graph/quiet-from.test.ts`: the synthetic tenant as it stood earlier.
