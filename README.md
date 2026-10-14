# Prism365

Prism365 is a community-built, privacy-conscious Microsoft 365 administration,
reporting, and auditing platform. It is an independent project: it is not affiliated
with, endorsed by, or derived from AdminDroid or Microsoft.

Public source and contribution issues: [nlj3/prism365](https://github.com/nlj3/prism365).

> **Status: pre-alpha pilot.** User investigation, server-side Microsoft sign-in, local AI,
> and private script import are implemented. Live Microsoft tenant acceptance is still required;
> production support and tenant write actions are not yet available. Demo mode is the default.

## Why this exists

Microsoft 365 administrators should be able to understand tenant health without buying
an opaque reporting suite or granting a tool more access than it needs. Prism365 aims
to provide:

- fast, accessible reporting for identity, audit, usage, and service health;
- transparent permission requests, with every permission tied to a feature;
- privacy-preserving defaults and tenant-controlled retention;
- explainable findings that cite the underlying evidence;
- safe administration workflows with preview, approval, and rollback where possible;
- a contributor experience that works well for humans and AI agents.

Prism365 does **not** treat activity logs, presence, message counts, or other telemetry
as proof of employee productivity or performance. Audit data is operational and security
evidence with known gaps—not a worker-scoring system.

## What runs today

- A responsive user investigation view, synthetic overview, and lifecycle preview.
- A Fastify API with session/consent, investigation, AI grant/analysis, and private package routes.
- Server-side MSAL authorization-code sign-in with PKCE, one-use state, nonce, verified ID tokens,
  an operator allowlist, HttpOnly sessions, Origin/CSRF protection, and rate limits.
- A real delegated adapter for selected user profile fields, license details, and up to 100
  sign-ins from a seven-day window. Optional sources fail independently with visible coverage.
- Shared runtime contracts using Zod.
- A Microsoft Graph HTTP boundary with host allowlisting, read-only-by-default behavior,
  pagination, `Retry-After` handling, and code-owned global/GCC/GCC High/DoD/China profiles.
- A machine-readable source and permission catalog that records cloud, role, license, and
  tenant-service prerequisites for each feature.
- A policy package that blocks unsupported productivity claims and classifies findings.
- A provider-neutral AI broker with trusted expiring grants, current actor authorization,
  bounded minimized evidence, supplied citations, and access-attempt/completion/failure records.
  An optional loopback Ollama integration analyzes directory/license/sign-in evidence.
- A preview-only lifecycle automation scaffold with reviewed onboarding/offboarding runbooks,
  a synthetic sample package, exact step/permission/recovery plans, and a private JSON script-import
  library with SHA256 content identity. Actual script execution remains unavailable.
- Encrypted PostgreSQL state, forced tenant row policies, a restricted runtime role, expiry cleanup,
  and an append-only application access ledger. Memory mode is limited to development.
- A non-root self-host container and optional PostgreSQL/local-model Compose profiles.
- Automated formatting, tests, type checks, builds, dependency updates, and CodeQL.

## Quick start

Requirements: Node.js 24 or newer and npm 11 or newer.

```bash
npm install
npm run dev:api
```

In a second terminal:

```bash
npm run dev:web
```

Open `http://localhost:5173`. The API listens only on `127.0.0.1:8787` by default.
See [Setup and pilot acceptance](docs/SETUP.md) to connect a non-production tenant, configure
durable storage, or run the self-host container. The core stack is React/Vite and Fastify;
there is no framework migration requirement to contribute.

Useful checks:

```bash
npm run check
npm run build
```

## Product boundaries

The MVP is intentionally narrow and read-only:

1. Tenant connection and consent inventory.
2. Entra users and sign-in/audit summaries.
3. Microsoft 365 usage report summaries.
4. Evidence-backed findings and export.
5. Optional, read-only AI explanations through an administrator-selected provider.
6. No management writes until preview, authorization, audit, and approval controls exist.

Purview Audit search is a separate optional connector because it has separate licensing,
role, query, retention, and permission constraints. It is not silently bundled into the
basic Graph connection.

The deployment target is any supported Microsoft Entra **work or school** tenant whose
administrator grants the selected permissions and whose user/service principal, licenses,
tenant settings, and Microsoft cloud satisfy that feature's prerequisites. Personal Microsoft
accounts are not supported. Features degrade independently rather than demanding one broad
permission bundle.

See [Data sources](docs/DATA_SOURCES.md), [Compatibility](docs/COMPATIBILITY.md),
[Deployment requirements](docs/REQUIREMENTS.md), [Roadmap](docs/ROADMAP.md),
[Permissions](docs/PERMISSIONS.md), [AI integration](docs/AI_INTEGRATION.md),
[Lifecycle automation](docs/AUTOMATION.md), and [Privacy principles](docs/PRIVACY.md) before
proposing a connector, report, or runbook.

## Architecture at a glance

```mermaid
flowchart LR
  Browser[Web dashboard] --> API[Application API]
  API --> Policy[Policy and authorization]
  API --> Jobs[Ingestion jobs]
  Jobs --> Graph[Microsoft Graph adapter]
  Jobs --> Purview[Optional Purview adapter]
  Jobs --> Activity[Optional Management Activity adapter]
  Jobs --> Workload[Isolated workload adapters]
  Graph --> M365[Customer tenant]
  Purview --> M365
  Activity --> M365
  Workload --> M365
  Jobs --> Store[(Tenant-isolated storage)]
  API --> Store
  API --> AI[Governed AI analysis broker]
  AI --> Store
  AI --> Model[Selected local or hosted model]
  AI --> Ledger
  API -. future approved plan .-> Runner[Isolated automation worker]
  Runner -. capability grant .-> M365
  Runner --> Ledger
  API --> Ledger[(Immutable action ledger)]
```

The current scaffold implements the browser, API contract, policy core, safe Graph boundary,
server-side identity, user investigation, encrypted state, governed local AI, and private script
imports. Scheduled ingestion, tenant-wide reporting, raw audit storage, additional AI providers,
and an isolated automation runner remain later milestones.
See [Architecture](docs/ARCHITECTURE.md) for the target design.

## Permission philosophy

- Start with no tenant permissions.
- Ask for a feature pack only when an administrator enables that feature.
- Prefer delegated or resource-specific consent when it satisfies the use case.
- Keep read and write identities separate.
- Never use `Directory.ReadWrite.All` or similarly broad permissions as a convenience.
- Show the exact permission, data accessed, purpose, and behavior if denied.
- Re-check Microsoft documentation during implementation; permissions and availability
  can change.

The catalog includes separate read packs for directory, license inventory, service health,
audit, usage, security, and device data. Purview, the Management Activity feed, and Exchange
Online automation remain distinct connectors with their own prerequisites.

## Contributing

Start with [CONTRIBUTING.md](CONTRIBUTING.md). Humans and AI agents use the same issue,
review, test, privacy, and attribution standards. AI agents must additionally follow
[AGENTS.md](AGENTS.md) and the coordination board in [TASKS.md](TASKS.md).

Pre-alpha ownership and decision rules are in [GOVERNANCE.md](GOVERNANCE.md). Repository
owners should follow the [publishing checklist](docs/PUBLISHING.md) before making a public
GitHub claim.

Good first contributions include accessibility fixes, mock-data stories, contract tests,
permission documentation, and small read-only report adapters. Please do not submit bulk
generated reports or tenant write actions without an accepted design issue.

## Security

Do not open a public issue for a vulnerability or include tenant data, access tokens,
client secrets, directory object identifiers, or audit records in an issue. Follow
[SECURITY.md](SECURITY.md).

## License and trademarks

Prism365 is licensed under the [Apache License 2.0](LICENSE). Apache-2.0 is permissive
and includes an explicit patent license, which is useful for a broad contributor base.

“Microsoft,” “Microsoft 365,” “Entra,” “Graph,” and “Purview” are trademarks of
Microsoft Corporation. “AdminDroid” belongs to its respective owner. These names are
used only to describe interoperability or project inspiration; no endorsement is implied.
