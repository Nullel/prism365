# Modules: service.health and service.messages

Status: built (TL-023). Package: `packages/module-service-health`.

## 1. What they answer

- **Is Microsoft down?** The incidents and advisories Microsoft has published for the services
  this tenant subscribes to, with Microsoft's own status, impact description, workaround and a
  link to the incident in the Microsoft 365 admin center.
- **What does Microsoft want me to do?** Message center posts with an action deadline, a
  retirement, a major planned change or a high severity.

Both feed the notification center (ADR-0007) and the Overview. They are tenant-wide, so they
are not shown on person pages: an advisory about a service someone uses says nothing about that
person.

## 2. Setup

Each module is independent and turns on when all of these are true:

| Module | Read pack | Graph permission (delegated, admin consent) | Signed-in role, for example |
|---|---|---|---|
| `service.health` | `service.health.read` | `ServiceHealth.Read.All` | Global Reader, Service Support Administrator |
| `service.messages` | `service.messages.read` | `ServiceMessage.Read.All` | Message Center Reader, Global Reader |

1. Add the permission to the app registration under **API permissions**, and grant admin consent.
2. Add the read pack to `ENABLED_READ_FEATURES`.
3. Sign in again, or use the **Grant access** link Prism365 shows, so your session carries it.
4. Start collecting. Service health is read every 5 minutes, Message center every 30.

A module that is not enabled is listed as "not set up" and raises nothing. One that is enabled but
not granted raises a "cannot be checked" alert with a grant link. Microsoft documents the roles
in its admin center guidance; the examples above are not a complete list.

## 3. Source

Microsoft Graph v1.0 service communications, checked 2026-10-22:
`admin/serviceAnnouncement/healthOverviews`, `admin/serviceAnnouncement/issues` and
`admin/serviceAnnouncement/messages`. Issues are read twice per run, open ones
(`isResolved eq false`) and those modified in the last 7 days. After a complete, fully readable
read of the open list, any incident Prism365 still shows as active that Microsoft no longer lists
as open is closed with the note "Microsoft no longer lists this as open.", so an incident resolved
while collection was stopped does not stay open. If a tenant rejects the filters, Prism365 reads
the whole list and filters it locally. Microsoft's enum values appear in more than one casing, so comparisons ignore case.

Post and message bodies are HTML. Prism365 stores and shows plain text only. From the newest
incident post it reads Microsoft's labeled sections: "Scope of impact" is shown with Microsoft's
impact description, and "Workaround" is shown only when Microsoft wrote one.

Admin center links use the pattern `https://admin.microsoft.com/#/servicehealth/:/alerts/{id}`
and `.../#/MessageCenter/:/messages/{id}`. The pattern is community-documented rather than a
published Microsoft specification, and the links open only for an administrator signed in to the
tenant.

## 4. The panel's states

The panel says whether the check itself worked before saying what it found:

| State | Meaning |
|---|---|
| `ok` | Checked within 15 minutes. No incidents reads "No issues reported by Microsoft". |
| `partial` | The latest check could not read every record, so an incident may be missing. |
| `stale` | The last successful check is older than 15 minutes (or collection is stopped). |
| `failed` | The latest attempt failed after the last success; the reason code is shown. |
| `never` | Enabled and granted, but no check has finished yet. |
| `missing_permission` | Enabled, but the session has not granted the permission. |
| `not_enabled` | The installation has not enabled the read pack. |

The panel never says a service is healthy. It shows Microsoft's impact text verbatim and does not
assume every user is affected.

## 5. What becomes a notification

Each incident and each qualifying post is one item, keyed by Microsoft's ID. Refreshing the same
facts changes nothing. These fields are "meaningful" and, when they change, add a revision, a
change note ("Status: Investigating to Service restored") and an unread mark:

- Incidents: status, classification, title, impact description, and resolution.
- Message center: title, action-required date, major-change flag, tags.

New posts on an unchanged incident update the text without notifying, so an incident Microsoft
updates hourly does not notify hourly.

Message center posts qualify by Microsoft's fields only: an action deadline no more than 7 days
past, or (modified in the last 90 days and) a Retirement tag, a major change in Plan for change, or
high or critical severity. Informational posts are stored but not shown. A post whose title only
suggests a retirement gets a separate Prism365 label, "Possible retirement (from the title)".

A post stops being active for one of two stated reasons: "No longer listed by Microsoft." when
Microsoft's feed no longer has it, or "No longer selected by Prism365: ..." when Prism365's own
rule (deadline long past, older than 90 days) drops it. Neither happens on an empty read or a read
with unreadable records, and a post that comes back keeps your review state.

Provider text is cleaned (control characters removed) and cut to size, never a reason to reject
a record; one record that still cannot be stored is counted as skipped and the run continues.

Severity in Prism365 is an ordering aid, set deterministically: an active incident is high
(critical for a service interruption), an advisory medium, a resolved item info; a post with a
deadline within 14 days is high, other deadlines and retirements medium, major changes low.

## 6. Retention

Provider records: issues 180 days, posts 365 days, the service list replaced each run.
Notification items Microsoft has closed or stopped listing: 90 days after last seen.

## 7. Tests

`packages/module-service-health/test`: impact, workaround and link; one item per incident across
refreshes; resolution notes; filter fallback; Message center labels, deadlines and the local
label; retirement of unlisted posts; a missing permission skipping only its module; every panel
state; HTML handling.
