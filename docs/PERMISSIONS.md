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

## Initial read-only packs

| Pack | Feature | Permission | Mode | Default | Data handling |
|---|---|---|---|---|---|
| `profile` | Signed-in profile | `User.Read` | Delegated | on at sign-in | session identity only |
| `directory.users.read` | User inventory | `User.Read.All` | Application for scheduled sync; delegated prototype allowed | off | normalized directory metadata; selected fields only |
| `audit.signins.read` | Sign-in and directory audit summaries | `AuditLog.Read.All` | Application for scheduled sync; delegated investigation allowed | off | normalized events/summaries with bounded retention |
| `reports.usage.read` | Microsoft 365 usage reports | `Reports.Read.All` | Application or delegated | off | aggregate trends; privacy setting respected |
| `purview.entra.read` | Optional Entra workload audit search | `AuditLogsQuery-Entra.Read.All` | Application or delegated | off/experimental | query results only under explicit retention policy |

`AuditLog.Read.All` is the documented least-privileged permission for listing sign-ins, and
`Reports.Read.All` is required for Office 365 usage reports. Delegated sign-in access also
depends on the signed-in user's Entra role. Conditional Access details require additional
policy-reading permission; TenantLens must show those details as unavailable rather than
requesting them silently.

Purview Audit Search is distinct from Entra directory audits. The query API has
workload-specific permissions and a broad `AuditLogsQuery.Read.All`; TenantLens starts with
the narrower Entra workload permission and does not request the broad permission by default.

## Not approved for the MVP

- Any `*.ReadWrite.*` permission.
- `Directory.ReadWrite.All` or `RoleManagement.ReadWrite.Directory`.
- Mail, chat, channel-message, file-content, calendar-content, or presence permissions.
- Broad Purview `AuditLogsQuery.Read.All` without a reviewed feature case.
- A permission added only because it makes an implementation simpler.

## Endpoint constraints

- Use `/v1.0` by default; beta endpoints require an ADR and a UI label.
- Use `$select` to limit fields and `$filter` only where the endpoint supports it.
- Follow `@odata.nextLink` as an opaque provider URL after validating its allowed host.
- Honor `429` and `Retry-After`; fall back to bounded exponential backoff only when needed.
- Prefer delta queries or change notifications over high-frequency polling where supported.
- Subscriptions have resource-dependent lifetimes and require renewal. Lifecycle events can
  report reauthorization, removal, or missed notifications, so schedule reconciliation.
- National-cloud support is a tested feature, not an assumed string replacement.

## Useful primary references

- [List sign-ins](https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0)
- [Office 365 active-user report](https://learn.microsoft.com/en-us/graph/api/reportroot-getoffice365activeuserdetail?view=graph-rest-1.0)
- [Create a Purview audit query](https://learn.microsoft.com/en-us/graph/api/security-auditcoreroot-post-auditlogqueries?view=graph-rest-1.0)
- [Graph throttling guidance](https://learn.microsoft.com/en-us/graph/throttling)
- [Change notifications](https://learn.microsoft.com/en-us/graph/change-notifications-overview)
- [Lifecycle notifications](https://learn.microsoft.com/en-us/graph/change-notifications-lifecycle-events)
