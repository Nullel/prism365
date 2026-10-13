# Data-source strategy

TenantLens is a federation of explicit connectors, not a screen scraper and not a single
oversized Microsoft Graph application. Every connector must declare its API, permissions,
tenant prerequisites, supported clouds, data sensitivity, retention, and failure states.

## Microsoft sources

| Source | Best use | Authentication and constraints | Initial decision |
|---|---|---|---|
| Microsoft Graph REST v1.0 | Directory, sign-ins, usage reports, service health, security, Intune, and selected configuration | Delegated or application permissions; delegated access is also bounded by the user's role; endpoint, license, and national-cloud availability vary | Primary MVP API, split into opt-in feature packs |
| Microsoft Purview Audit Search through Graph | Bounded investigations across enabled audit workloads | Separate workload-specific permissions, roles, licenses, retention, asynchronous queries, and service limits | Optional connector; disabled by default |
| Office 365 Management Activity API | Continuous audit ingestion from subscribed content blobs | `ActivityFeed.Read`, unified audit logging, tenant-matched tokens, subscription lifecycle, cloud-specific service roots | Durable-ingestion source after the Graph MVP |
| Exchange Online PowerShell | Exchange configuration not exposed by stable Graph APIs | Certificate or federated app identity, `Exchange.ManageAsApp`, and narrowly scoped Exchange RBAC | Isolated optional worker; never run in the web process |
| SharePoint admin APIs | Tenant and site administration gaps across Graph, REST, and CSOM | Application permission plus SharePoint/Entra authorization; API choice depends on operation | Add only for a reviewed report; do not use retired Azure ACS app-only auth |
| Microsoft Graph Data Connect | High-volume bulk extraction | Azure/Fabric data pipeline, admin consent, storage, billing, and no national-cloud support | Optional enterprise adapter, not an MVP dependency |
| Teams resource-specific consent | Data scoped to a team/chat where a Teams app is installed | Teams app package and resource-owner consent; not a general tenant-reporting permission | Consider only for genuinely resource-scoped features |

There is no universal Microsoft 365 reporting permission. Each row above is a distinct
security and operational boundary. A connector must not silently fall back to another source
with broader access.

### API behavior we must preserve

- Treat Microsoft Graph `@odata.nextLink` as opaque, but reject links outside the selected
  cloud's allowlisted Graph host.
- Honor `Retry-After`, bound pagination and retries, and use delta/change notifications when
  the endpoint supports them.
- Record source time, collection time, coverage, pagination completion, and partial-result
  warnings with every derived finding.
- Keep the tenant's usage-report concealment setting. If it cannot be read, default to the
  more private presentation.
- Do not infer that missing audit records prove inactivity, or that activity volume measures
  productivity, performance, effort, or attendance.

## Open-source projects we can learn from

| Project | Useful contribution to TenantLens | License and integration approach |
|---|---|---|
| [CISA ScubaGear](https://github.com/cisagov/ScubaGear) | Machine-testable Microsoft 365 security-baseline ideas and evidence-oriented assessment patterns | CC0-1.0. Consume versioned public baselines with source/version provenance; do not present CISA guidance as a TenantLens guarantee |
| [Maester](https://github.com/maester365/maester) | Extensible Microsoft 365 security tests and result-export patterns | MIT. Prefer an optional import/adapter or independently implemented compatible evidence schema |
| [Microsoft365DSC](https://github.com/Microsoft365DSC/Microsoft365DSC) | Configuration extraction, drift concepts, and broad workload coverage | MIT. Treat exported snapshots as an optional source; do not make PowerShell a requirement for the core UI |

Using a public repository as research does not mean copying its code. Before importing code,
rules, schemas, or assets, record the exact version, license, attribution, update path, and
security review. Add third-party notices when required.

## Explicit exclusions

- No scraping Microsoft admin-center pages or relying on private browser endpoints.
- No reverse engineering or copying AdminDroid reports, names, layouts, assets, or code.
- No deprecated Azure ACS app-only authentication for SharePoint.
- No mailbox, message, file-content, calendar-content, or presence collection merely because
  an API makes it possible.
- No Graph beta dependency without an architecture decision, fallback, and visible beta label.

## Primary references

- [Microsoft Graph permissions overview](https://learn.microsoft.com/en-us/graph/permissions-overview)
- [Microsoft Graph national cloud deployments](https://learn.microsoft.com/en-us/graph/deployments)
- [Office 365 Management Activity API](https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference)
- [Microsoft Purview Audit query API](https://learn.microsoft.com/en-us/graph/api/security-auditcoreroot-post-auditlogqueries?view=graph-rest-1.0)
- [Exchange Online PowerShell app-only authentication](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2)
- [SharePoint admin API authorization](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/sharepoint-admin-apis-authentication-and-authorization)
- [SharePoint app-only retirement guidance](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly)
- [Microsoft Graph Data Connect FAQ](https://learn.microsoft.com/en-us/graph/data-connect-faq)
- [Teams resource-specific consent](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/app-permissions/teams-app-permissions)
