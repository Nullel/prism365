# ADR-0007: notification center ("attention")

- Status: accepted; implemented as TL-023
- Date: 2026-10-22
- Related: ADR-0006 (module runtime), `docs/modules/service-health.md`

## Context

Prism365 needs one place for things an administrator may have to act on: Microsoft incidents,
Message center deadlines, and later expiring secrets, license capacity, breach exposure and device
compliance. Each source changes on its own schedule, and the same facts arrive on every refresh.
Findings (ADR-0006) are Prism365's own detections with an evidence model; provider announcements
are not detections and need different change rules.

## Decision

1. **One store, `prism365_attention`, written only through `context.attention`.** Connectors call
   `upsert` with the provider's facts and a small `meaningful` map, and `retire` after a complete
   read of one kind. Items are keyed by module, kind and the provider's ID.
2. **Change is decided deterministically.** A change to a meaningful field, a resolution, or a
   reactivation adds a revision and a plain change note; anything else refreshes text silently. No
   model is involved. Optional AI summaries, if added, must keep the original link and evidence.
3. **Provider facts and administrator state never share a column.** Microsoft's status, labels,
   deadline and active flag are separate from unread, review state (new, reviewed, done) and
   muted. Reviewing never changes what Microsoft said, and Microsoft resolving an item never marks
   it reviewed. An item that becomes active again returns to "new".
4. **Labels Prism365 infers are stored apart** (`local_labels`), so they are never shown as the
   provider's.
5. **Notification is a preference, not a source.** `prism365_preference` holds per-tenant settings:
   important services (none selected means all) and a switch per class of item. Muted items never
   become unread.
6. **Monitoring problems are computed, not stored.** "Not collecting", "cannot be checked" and
   "check failed" come from collection state on each request, so they cannot go stale.
7. **Only implemented, enabled sources appear.** A source the installation never enabled is "not
   set up" and raises nothing.

## Consequences

- New attention categories are new modules; the store, inbox and settings need no change. An
  older build shows a class it does not know as "Other items" and keeps a newer build's
  notification switches when it saves settings.
- A source's availability is judged by the credential that collects while collection runs, and
  by the viewer's session only when it is stopped.
- Delivery is in-app only. Teams, Slack or email delivery would read the same unread transitions.
- Items closed or dropped by the provider are deleted 90 days after last seen.
