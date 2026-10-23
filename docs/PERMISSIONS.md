# Microsoft permissions policy

This file is a design inventory, not a substitute for the current Microsoft endpoint
documentation. Re-verify a permission when implementing or releasing the feature.

## Consent rules

1. No “grant everything” setup path.
2. Permissions are grouped by visible feature, not product edition.
3. The consent screen explains the data accessed, purpose, storage, and behavior if denied.
4. Prefer delegated access for interactive work when it satisfies the scenario. Background
   collection uses application access only when necessary and after admin consent.
5. Read and future write permissions use separate application identities.
6. An unavailable permission disables only its dependent feature.
7. Permission changes require a design issue, this inventory update, and security review.

Microsoft recommends requesting the fewest permissions required and documents that
application permissions carry greater privacy risk because they work without a signed-in
user. See [Microsoft Graph permission best practices](https://learn.microsoft.com/en-us/graph/best-practices-graph-permission).

## Read-only feature packs

The executable catalog in `packages/graph/src/permissions.ts` is the detailed source of truth.
This table groups the intended scope for review.

| Packs | Source | Permission(s) | Phase | Important prerequisite |
|---|---|---|---|---|
| `profile`, `directory.users.read` | Graph | `User.Read`, `User.Read.All` | MVP | Work/school tenant; delegated access is bounded by the user's privileges |
| `tenant.licenses.read` | Graph | `LicenseAssignment.Read.All` | MVP | Delegated caller needs an allowed directory role |
| `user.licenses.read` | Graph | `LicenseAssignment.Read.All` | Pilot implemented | `/users/{id}/licenseDetails` supports delegated access only, plus an allowed Entra role |
| `service.health.read` | Graph | `ServiceHealth.Read.All` | MVP | Service health available for tenant services |
| `audit.signins.read` | Graph | `AuditLog.Read.All` | MVP | Sign-ins only; delegated caller needs an allowed role; license affects fields/retention |
| `audit.directory.read` | Graph | `AuditLog.Read.All` | MVP | Directory audit log for the change feed (`directory.changes`); delegated caller needs Reports Reader, Security Reader or Security Administrator (Learn's list); Learn disagrees on whether Graph access needs Entra ID P1/P2; Microsoft keeps 7 days (30 with P1/P2); Prism365 keeps what it collects with no retention limit |
| `directory.applications.read` | Graph | `Application.Read.All` | Experimental | App consents audit: enterprise applications and the application permissions granted to them; Microsoft does not list Global Reader for reading app role assignments |
| `directory.permission-grants.read` | Graph | `Directory.Read.All` | Experimental | Optional for the app consents audit: delegated consents. `Directory.Read.All` is Microsoft's documented least privilege for listing grants (both the tenant-wide and the per-application list), and it can read the whole directory, so enable it only if you want delegated consents shown |
| `reports.authentication-methods.read` | Graph | `AuditLog.Read.All` | MVP | MFA registration audit; Reports Reader, Security Reader, Security Administrator or Global Reader; excludes disabled accounts |
| `reports.usage.read`, `reports.settings.read` | Graph | `Reports.Read.All`, `ReportSettings.Read.All` | MVP | Only the per-user Microsoft 365 apps report (`getM365AppUserDetail`, JSON) and the name-concealment setting; per-user detail needs Reports Reader or a service admin role; global/GCC target only |
| `identity.conditional-access.read` | Graph | `Policy.Read.All` | Experimental | Conditional Access policies audit; Global Reader, Security Reader, Security Administrator or Conditional Access Administrator; Entra ID P1; 1 request per second per tenant, shared with risky users |
| `identity.risk.read` | Graph | `IdentityRiskyUser.Read.All` | Experimental | Risky users audit; Entra ID P2; Global Reader, Security Reader, Security Operator or Security Administrator; 1 request per second per tenant, shared with Conditional Access; not targeted for China |
| `security.score.read` | Graph | `SecurityEvents.Read.All` | Experimental | Applicable subscriptions; not targeted for China |
| `security.incidents.read` | Graph | `SecurityIncident.Read.All` | Experimental | Defender entitlement and security role; not targeted for China |
| `devices.intune.read` | Graph | `DeviceManagementManagedDevices.Read.All` | Experimental | Active Intune license and role |
| `purview.entra.read` | Graph/Purview | `AuditLogsQuery-Entra.Read.All` | Experimental | Purview Audit availability, roles, query limits, and retention |
| `audit.activity-feed.read` | Management Activity | `ActivityFeed.Read` | Future | Unified audit enabled and content subscription started |
| `exchange.configuration.read` | Exchange Online PowerShell | `Exchange.ManageAsApp` | Future | Certificate/federation plus narrow Exchange RBAC |

`AuditLog.Read.All` is the documented least-privileged permission for listing sign-ins, and
`Reports.Read.All` is required for Office 365 usage reports. Delegated sign-in access also
depends on the signed-in user's Entra role. Conditional Access details require additional
policy-reading permission; Prism365 must show those details as unavailable rather than
requesting them silently.

Purview Audit Search is distinct from Entra directory audits. The query API has
workload-specific permissions and a broad `AuditLogsQuery.Read.All`; Prism365 starts with
the narrower Entra workload permission and does not request the broad permission by default.

Having the API permission is not, by itself, proof that a feature can run. Prism365 also
needs the signed-in user or service principal role, product license, tenant configuration,
endpoint/cloud availability, and provider health. The pilot validates configured cloud, actor,
and granted scopes, then reports source-specific failures; it does not yet implement exhaustive
role/license/service probes. A provider 403 is classified as uncertain, not automatically missing
consent. See [Compatibility](COMPATIBILITY.md).

## Not approved for the MVP

- Any `*.ReadWrite.*` permission.
- `Directory.ReadWrite.All` or `RoleManagement.ReadWrite.Directory`.
- Mail, chat, channel-message, file-content, calendar-content, or presence permissions.
- Broad Purview `AuditLogsQuery.Read.All` without a reviewed feature case.
- A permission added only because it makes an implementation simpler.

## Future lifecycle action packs — disabled

The preview contracts name these permissions so an administrator can understand the eventual
impact. They are not requested by the current app, are not present in the executable read catalog,
and cannot be used because no execution route or write identity exists.

| Capability | Candidate least permission / authorization | Constraint before approval |
|---|---|---|
| Create a user | `User.Create` | Approved property template; separate write identity |
| Disable or enable a user | `User.EnableDisableAccount.All` + `User.Read.All` | Entra role and target-role checks; protected accounts blocked |
| Revoke sessions | `User.RevokeSessions.All` | Eventual propagation; external-user limitation disclosed |
| Assign or remove licenses | `LicenseAssignment.ReadWrite.All` | Capacity, dependency, and retention preview |
| Add or remove static group members | `GroupMember.ReadWrite.All` | Dynamic and role-assignable groups excluded by default |
| Mailbox handoff | Resource-scoped Exchange Online RBAC | Separate connector/identity; recipient and resource scope |
| File/data handoff | Workload-specific resource authorization | Ownership, sharing, retention, hold, and region policy |

Prism365 will not replace these packs with `Directory.ReadWrite.All`. Authentication-method
deletion, user deletion, role assignment, dynamic-group mutation, and legal-hold changes require
separate designs and are not part of the initial lifecycle runbooks. See
[Lifecycle automation](AUTOMATION.md).

## Endpoint constraints

- Use `/v1.0` by default; beta endpoints require an ADR and a UI label.
- Use `$select` to limit fields and `$filter` only where the endpoint supports it.
- Follow `@odata.nextLink` as an opaque provider URL after validating its allowed host.
- Honor `429` and `Retry-After`; fall back to bounded exponential backoff only when needed.
- Prefer delta queries or change notifications over high-frequency polling where supported.
- Subscriptions have resource-dependent lifetimes and require renewal. Lifecycle events can
  report reauthorization, removal, or missed notifications, so schedule reconciliation.
- National-cloud support is a tested feature, not an assumed string replacement.
- The selected cloud fixes the authority, Graph host, and workload endpoints; tenant-provided
  URLs never override that allowlist.

## Useful primary references

- [List sign-ins](https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0)
- [Office 365 active-user report](https://learn.microsoft.com/en-us/graph/api/reportroot-getoffice365activeuserdetail?view=graph-rest-1.0)
- [Create a Purview audit query](https://learn.microsoft.com/en-us/graph/api/security-auditcoreroot-post-auditlogqueries?view=graph-rest-1.0)
- [Graph throttling guidance](https://learn.microsoft.com/en-us/graph/throttling)
- [Change notifications](https://learn.microsoft.com/en-us/graph/change-notifications-overview)
- [Lifecycle notifications](https://learn.microsoft.com/en-us/graph/change-notifications-lifecycle-events)
- [List subscribed SKUs](https://learn.microsoft.com/en-us/graph/api/subscribedsku-list?view=graph-rest-1.0)
- [Service health overview](https://learn.microsoft.com/en-us/graph/api/serviceannouncement-list-healthoverviews?view=graph-rest-1.0)
- [Usage-report authorization](https://learn.microsoft.com/en-us/graph/reportroot-authorization)
- [Read report privacy settings](https://learn.microsoft.com/en-us/graph/api/adminreportsettings-get?view=graph-rest-1.0)
- [List Conditional Access policies](https://learn.microsoft.com/en-us/graph/api/conditionalaccessroot-list-policies?view=graph-rest-1.0)
- [Get a risky user](https://learn.microsoft.com/en-us/graph/api/riskyuser-get?view=graph-rest-1.0)
- [List Secure Scores](https://learn.microsoft.com/en-us/graph/api/security-list-securescores?view=graph-rest-1.0)
- [List Defender incidents](https://learn.microsoft.com/en-us/graph/api/security-list-incidents?view=graph-rest-1.0)
- [List Intune managed devices](https://learn.microsoft.com/en-us/graph/api/intune-devices-manageddevice-list?view=graph-rest-1.0)
- [Create user](https://learn.microsoft.com/en-us/graph/api/user-post-users?view=graph-rest-1.0)
- [Update user](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0)
- [Revoke sign-in sessions](https://learn.microsoft.com/en-us/graph/api/user-revokesigninsessions?view=graph-rest-1.0)
- [Assign license](https://learn.microsoft.com/en-us/graph/api/user-assignlicense?view=graph-rest-1.0)
- [Add group member](https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0)
- [Exchange Online application RBAC](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac)

## Packs share permissions; Prism365 enforces packs

Several packs carry the same Microsoft permission (`AuditLog.Read.All` serves sign-ins, the
directory audit log and MFA registration; `LicenseAssignment.Read.All` serves per-user and tenant
licences). Microsoft grants permissions, not packs, so consenting one of them grants all. Prism365
therefore enforces the pack itself: a module collects, an audit runs, and a route reads only when
every pack it uses is in `ENABLED_READ_FEATURES`, whatever the session's token would allow.

Rolling back to a build older than 2026-10-23: remove `audit.directory.read`,
`identity.conditional-access.read`, `identity.risk.read`, `devices.intune.read`,
`directory.applications.read`, `directory.permission-grants.read`, `reports.usage.read` and
`reports.settings.read` from `ENABLED_READ_FEATURES` first. Older builds refuse to start when the
list names a feature they do not know.

