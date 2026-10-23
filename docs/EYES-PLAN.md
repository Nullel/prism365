# The eyes: what Prism365 sees next

Decided 2026-10-23. Prism365 is the eyes: it reads a tenant and never writes to it (see
`ROADMAP.md`, Milestone 6). Anything that acts on a tenant, such as offboarding, belongs to a
separate action tool. The two meet through files, never through shared credentials: Prism365
exports a list, the action tool sends back a receipt.

This plan covers the next set of read-only features. Each is a slice with its own branch, risk
class and review. RED slices (stored data, background use of a session, certificates, anything
written to Microsoft by the setup script) are built by the strongest model and get an adversarial
review before they are committed. YELLOW slices are built from a written contract by the
contract implementer and reviewed.

## Groundwork (done)

New read packs (`audit.directory.read`, `directory.applications.read`,
`directory.permission-grants.read`), every planned pack selectable in `ENABLED_READ_FEATURES`,
a route file per feature under `apps/api/src/routes`, a contracts file per feature, page slots
in the navigation and on the person page, and a file per feature in the demo's synthetic Graph.

## Slices

| Slice | What it gives | Class | New permission | Depends on |
|---|---|---|---|---|
| S1 Directory change feed | Every directory audit event kept for as long as Prism365 runs (Microsoft keeps 7 or 30 days), a filterable feed, changes on each person's page, findings for risky changes | RED | none (`AuditLog.Read.All`) | groundwork |
| S2 Audit trends and schedules | Each audit's tiles over time, kept forever; audits that run daily or weekly on the collecting session | RED | none | groundwork |
| S3 This week | One page for the last seven days: new findings, accounts gone quiet, licence and MFA drift, directory changes, Microsoft deadlines | YELLOW | none | S1, S2 |
| S4 Setup | `npm run setup` registers the app, makes and uploads the certificate and writes `.env`; a first-run checklist on the Microsoft access page | RED | none at runtime; the script writes the app registration | groundwork |
| S5 Handoffs | Export an audit's list for an action tool; import the receipts it sends back; show them on the person page | RED | none | groundwork |
| S6 Consent-batch audits | Conditional Access, risky users, devices, app consents, Microsoft 365 app usage | YELLOW | one batched request: `Policy.Read.All`, `IdentityRiskyUser.Read.All`, `DeviceManagementManagedDevices.Read.All`, `Application.Read.All`, `DelegatedPermissionGrant.Read.All`, `Reports.Read.All`, `ReportSettings.Read.All` | groundwork |
| S7 Mail forwarding | Mailboxes and inbox rules that forward mail | RED | decided by research: Graph delegated access cannot read other mailboxes' settings | research |

## Order

Wave 1 builds S1, S2, S4, S5 and S6 side by side, each in its own worktree. Wave 2 builds S3 on
top of S1 and S2, and S7 if the research finds a read path that keeps Prism365 read-only. After
the waves: the demo is re-recorded once, every new screen is checked in the browser, the docs are
brought up to date, and the branch goes up as one pull request.

## What stays out

No write path in the runtime, no action buttons, no shared credentials with the action tool.
The setup script is the only code in the repository that writes to Microsoft, and it is a tool
the administrator runs once, on purpose, after reading its plan.
