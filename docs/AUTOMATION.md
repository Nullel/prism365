# Lifecycle automation and imported runbooks

Prism365 will let an authorized administrator choose onboarding or offboarding, enter a work
email, select a reviewed runbook, preview the exact impact, approve it, and follow the result in
the UI. The fast interaction must not hide the fact that these workflows can create, disable, or
remove access across an organization.

## Current implementation

The pre-alpha includes:

- runtime-validated manifest, workflow, and preview contracts;
- one native onboarding sample and one synthetic imported PowerShell 7 offboarding package;
- a preview-only API and UI using sample workflow metadata, with steps, risk, recovery, and
  permission requirements; these previews do not inspect the target tenant;
- authenticated, CSRF-protected JSON package imports into encrypted tenant-scoped PostgreSQL
  storage, plus a metadata-only listing API and UI;
- structural inspection, content SHA256, immutable imported records, and `unreviewed` status;
- no review/approval route, tenant write connection, script process, execution endpoint, or write
  permission.

The imported sample is metadata only. It demonstrates how an organization-owned script will
appear after review; no PowerShell file is included or executed.

Actual imports contain caller-owned source and remain private. They require a signed-in live tenant
and durable encrypted storage; they are unavailable in demo mode. Source packages expire after
30 days and their listing returns only metadata. The first live-auth flow remains an unverified
pilot; an authorized tenant test is required before claiming production compatibility.

## Target execution flow

Import and structural inspection are implemented. Review, approval, real-state preview, and all
execution steps below remain future work.

1. **Import:** an administrator uploads a versioned JSON automation package containing a manifest
   and text files. ZIP/archive extraction is not implemented.
2. **Inspect:** Prism365 validates paths, sizes, the manifest and fixed entrypoint, computes a digest,
   and leaves the package unreviewed. Static/security content scanning and signature review remain
   required before a future approval flow; structural inspection does not establish trust.
3. **Review:** a separate authorized reviewer approves that exact immutable digest for selected
   tenants and parameters. A changed byte creates a new unreviewed version.
4. **Preview:** Prism365 resolves the target, protected-account rules, current access, before-state,
   expected changes, required permissions, and recovery options without writing.
5. **Confirm:** the operator completes step-up authentication and types a target-specific phrase.
   High-impact policy may require a second approver.
6. **Queue:** the API records immutable intent and sends only an authorized capability grant to an
   isolated worker. It does not start a script process itself.
7. **Execute:** the worker applies one bounded step at a time, recording provider request IDs and
   before/after evidence. Uncertain results stop for reconciliation instead of blind retry.
8. **Reconcile:** Prism365 reads the resulting tenant state, closes successful steps, and exposes
   recovery actions or manual follow-up for partial failure.

## Automation package contract

The current package is JSON containing a manifest, the fixed entrypoint's text file, and optional
non-secret supporting text files. `POST /v1/automation/packages` imports it and
`GET /v1/automation/packages` returns metadata summaries. The manifest and script text remain
untrusted data; importing them does not execute or approve them.

```json
{
  "manifest": {
    "schemaVersion": 1,
    "id": "example.secure-offboarding",
    "name": "Secure employee offboarding",
    "version": "1.2.0",
    "description": "Organization-owned draft awaiting review.",
    "lifecycle": "offboarding",
    "runtime": "powershell7",
    "entrypoint": "scripts/offboard.ps1",
    "publisher": "Example organization",
    "parameters": [],
    "capabilities": ["identity.lookup"],
    "allowedHosts": [],
    "timeoutSeconds": 600,
    "steps": [{
      "id": "verify-target",
      "label": "Verify the target",
      "description": "Inspect account state before any future action.",
      "capability": "identity.lookup",
      "risk": "readOnly",
      "reversible": true,
      "recovery": "No tenant change is made."
    }]
  },
  "files": [{
    "path": "scripts/offboard.ps1",
    "content": "param([string] $TargetEmail)\n# No tenant changes in this sample.\n"
  }]
}
```

The manifest schema is in `packages/contracts`; package inspection and metadata views are in
`packages/automation/src/registry.ts`. Important rules include:

- PowerShell 7 only, with a present `.ps1` entrypoint;
- fixed relative paths; no traversal, absolute paths, control characters, duplicate paths or
  case-colliding paths;
- between one and 50 text files, with a 500000-byte cap covering canonical manifest and file
  content; binary NUL-containing source is rejected;
- empty `allowedHosts`, reserving the capability-broker mode without direct network access;
- SHA256 over canonical manifest and sorted file paths with exact source bytes; changed source,
  manifest, or line endings changes the digest and remains unreviewed;
- tenant, importer, digest, time and trust state assigned by the server, never the upload;
- no secret parameters in manifests or browser-submitted values;
- declared, unique capabilities and steps;
- each step declares impact, reversibility, and a recovery path;
- no arbitrary tenant-supplied executable path, command-line flags, Graph URL, or identity host;
- no wildcard future capability and no silent permission expansion.

## Execution boundary

`packages/automation` inspects imported source as data and compiles sample workflow metadata into
plans; it is not a script host. Imported packages cannot become runnable through their manifest or
upload fields, and no source-download or execution route is provided by this implementation.
Production execution requires a separate short-lived worker with all of these controls:

- a fresh isolated container or equivalent sandbox for one tenant and run;
- a read-only package mount identified by an approved cryptographic digest;
- non-root execution, read-only base image, bounded CPU/memory/process count/time, and no host
  socket or host filesystem access;
- denied outbound network by default, with code-owned Microsoft cloud endpoints allowlisted only
  for a reviewed direct-access compatibility mode;
- secrets delivered out of band, never as script parameters, files, environment dumps, or logs;
- tenant- and run-scoped authorization, idempotency, cancellation, and output limits;
- destruction of the sandbox and ephemeral storage after the result is durably recorded;
- sanitized structured events rather than raw terminal output in the browser.

The preferred mode is a **capability broker**: the script requests a named action and Prism365
performs the provider call after policy checks. The script never receives a Microsoft access token.
A legacy direct-access mode, if ever accepted, requires a separate ADR and a dedicated constrained
identity because a token exposed to arbitrary code cannot be reliably contained by API-level rules.

Local development of a future container runner must use the supported sandbox runtime (for example,
Docker) rather than executing imported scripts on the developer host.

## Least-privilege action packs

These are future packs, disabled by default and not approved for the current MVP. Permission is only
one prerequisite; delegated/admin role, product license, cloud, tenant policy, target role, and
resource state can independently block a step.

| Capability | Microsoft surface | Candidate least permission / authorization | Important constraint |
|---|---|---|---|
| Create account | Graph `POST /users` | `User.Create` | Separate write identity and approved property profile |
| Disable account | Graph `PATCH /users/{id}` | `User.EnableDisableAccount.All` + `User.Read.All` | Targeting administrators has extra Entra role constraints |
| Revoke sessions | Graph `POST /users/{id}/revokeSignInSessions` | `User.RevokeSessions.All` | Revocation can take several minutes; guest home tenant controls external sessions |
| Assign/remove license | Graph license assignment | `LicenseAssignment.ReadWrite.All` | Check dependencies, capacity, and retention before removal |
| Add/remove group membership | Graph group members | `GroupMember.ReadWrite.All` | Dynamic groups are not manually changed; role-assignable groups are excluded by default |
| Mailbox handoff | Exchange Online | Resource-scoped Exchange RBAC | Keep separate from Graph consent and scope to allowed recipients/resources |
| File/data handoff | Workload-specific | Resource-specific authorization | Retention, legal hold, ownership, sharing, and region require separate policy |

Prism365 will not request `Directory.ReadWrite.All` as a shortcut. Role-assignable group changes
would additionally require powerful role-management authority and are therefore outside the default
lifecycle capability.

Primary references, re-verified before implementation:

- [Create user](https://learn.microsoft.com/en-us/graph/api/user-post-users?view=graph-rest-1.0)
- [Update user](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0)
- [Revoke sign-in sessions](https://learn.microsoft.com/en-us/graph/api/user-revokesigninsessions?view=graph-rest-1.0)
- [Assign license](https://learn.microsoft.com/en-us/graph/api/user-assignlicense?view=graph-rest-1.0)
- [Add group member](https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0)
- [Exchange Online application RBAC](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac)

## Protected targets and policy stops

Execution must stop before mutation when the target is:

- the acting operator;
- a configured break-glass/emergency account;
- a privileged or protected administrator outside the approver's allowed target class;
- a guest/external account whose home tenant makes the requested outcome impossible;
- outside the selected tenant or verified work/school domains;
- already in an uncertain or conflicting lifecycle run.

Password reset, authentication-method deletion, permanent user deletion, role assignment, dynamic
group changes, litigation/retention changes, and mailbox deletion are not part of the initial
runbooks.

## AI boundary

The current AI integration analyzes read-only user-investigation datasets and has no imported
script access. A future provider may explain a preview, summarize failures, or draft a manifest from an
administrator's requirements. It cannot import, approve, alter, schedule, or execute a package and
never receives a tenant write credential. AI output is untrusted input that must pass the same
schema, review, authorization, preview, and confirmation path as human-authored content.

## Remaining implementation gates

Before imported code can run, implement digest-bound reviewer approvals, actual tenant-state
previews, protected-account policy, separate write identity, durable action intent, idempotency,
step-up confirmation, isolated runner, capability broker, cancellation/kill switch, reconciliation,
and partial-failure recovery. Test against an explicitly authorized tenant and obtain independent
security review. Importing a script is not an indication that its declared actions are authorized,
its code is safe, or its requested outcome is possible.

## Portability

Prism365 targets Microsoft Entra work or school tenants. A workflow works only when that tenant's
administrator has explicitly enabled its action packs and the caller/worker also satisfies the
required Entra or workload role, licenses, tenant settings, cloud support, and target restrictions.
The UI must show which prerequisite is missing instead of claiming that consent alone is enough.
