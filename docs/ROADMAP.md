# Roadmap

Dates are intentionally omitted until maintainers know contributor capacity. Milestones are
outcome gates, not feature promises.

## Milestone 0 — trustworthy foundation

Goal: a contributor can run, understand, test, and safely extend the project.

- [x] Original project identity, Apache-2.0 license, contribution and agent guidance.
- [x] Responsive, accessible demo dashboard and read-only API.
- [x] Shared contracts, Graph boundary, permission catalog, and policy tests.
- [x] Source inventory, Microsoft-cloud profiles, and compatibility contract.
- [x] Provider-neutral AI broker, access-grant contract, and synthetic isolation/policy tests.
- [x] CI, dependency updates, CodeQL, issue forms, and security policy.
- [ ] Publish the repository after owner authentication and settings review.
- [ ] Enable branch protection, private vulnerability reporting, secret scanning, and
  required CI checks in GitHub.
- [ ] Record maintainer and release governance.

Exit: the default branch is public, reproducible, green, and protected; no production
tenant connection is advertised.

## Milestone 1 — connect one test tenant, read-only

Goal: an administrator can understand and grant only the read scopes for selected features.

- Microsoft identity sign-in and tenant onboarding.
- Bring-your-own single-tenant app registration for the first self-host release.
- Cloud selection/discovery with separate deployment registrations per national cloud.
- Credential handling design and threat-model review.
- Consent inventory and per-feature readiness screen covering permission, role, license,
  tenant setting, provider health, implementation status, and cloud support.
- Read-only user inventory (`User.Read.All`).
- Read-only license inventory (`LicenseAssignment.Read.All`) and service health
  (`ServiceHealth.Read.All`) as independent packs.
- Read-only directory/sign-in audit summary (`AuditLog.Read.All`).
- Read-only Microsoft 365 usage summary (`Reports.Read.All`).
- Clear errors for missing role, license, permission, data, or unsupported cloud.
- Fully synthetic integration fixtures and an opt-in test-tenant runbook.
- Compatibility tests for Worldwide and any claimed GCC/GCC High/DoD/China feature; personal
  Microsoft accounts are explicitly unsupported.

Exit: no write scopes; consent can be removed cleanly; tokens and tenant records stay out of
logs; a security reviewer approves the connector boundary.

## Milestone 2 — durable ingestion and evidence-backed findings

Goal: reports remain reliable across throttling, pagination, restarts, and data gaps.

- PostgreSQL schema with row-level tenant isolation and deletion workflow.
- Per-tenant job queue, bounded retry, cursor/delta tracking, and dead-letter visibility.
- Data freshness, coverage, and last-success indicators on every report.
- Deterministic finding rules with source, time window, rule version, and limitations.
- Configurable retention and export/deletion controls.
- Cross-tenant isolation, throttling, pagination, and recovery test suites.

Exit: restart-safe ingestion; recovery runbooks; privacy review; restore test; no individual
productivity scoring.

## Milestone 3 — administrator reporting MVP

Goal: the project is useful for routine tenant review without management writes.

- Saved filters and reports.
- Scheduled exports with approved destinations and redaction options.
- Role-based dashboards for global, security, reports, and help-desk readers.
- Entra hygiene and sign-in investigation views.
- Usage adoption views with explicit interpretation warnings.
- Accessibility audit, performance budget, localization foundation, and mobile review.

Exit: documented self-host install and upgrade; backup/restore; release artifacts and SBOM;
pilot feedback from administrators.

## Milestone 4 — optional read-only AI analyst

Goal: let an administrator connect a chosen model to authorized Prism365 data without giving
the model tenant credentials, silent future access, or administrative tools.

- Local, tenant-hosted, and external provider adapter manifests.
- Provider setup screen with data use, training, retention, region, cost, and secret status.
- Exact expiring dataset grants; identifiers, raw records, and external processing off by default.
- Tenant-scoped bounded retrieval with source/time/coverage evidence and supplied-ID citations.
- Prompt-injection, cross-tenant, hallucinated-citation, privacy, cost, and model-upgrade evals.
- AI labels, limitations, answer feedback, delete controls, and emergency provider kill switch.
- No productivity inference and no action/write tools.

Exit: one local or tenant-hosted adapter passes conformance and an authorized test-tenant pilot;
disabling it stops new disclosure and removes configured credentials cleanly.

## Milestone 5 — optional Purview Audit connector

Goal: query Purview audit records without implying universal availability or completeness.

- Separate connector and consent flow.
- Workload-specific least-privilege packs where feasible.
- Query status, service limits, record-limit warnings, and partial-result handling.
- Licensing/role/retention discovery and clear UI status.
- Explicit raw-record retention controls and sensitive-field review.

Exit: connector works in an authorized test tenant and is disabled by default.

## Milestone 6 — guarded administration actions

Goal: introduce a very small number of reversible or recoverable actions safely.

- Separate write application identity and permission review.
- Dry-run diff, policy engine, step-up confirmation, idempotency, ledger, reconciliation.
- First action selected through a threat-model and user-research issue.
- Break-glass disable switch and incident runbook.

Exit: independent security review and no broad “manage everything” permission bundle.

## Explicitly out of scope for the first release

- On-premises Active Directory agents.
- An AI agent with autonomous tenant write privileges.
- Employee productivity, engagement, or performance scoring.
- Full commercial-product feature parity.
- Indefinite raw audit-log warehousing by default.
- Beta Graph endpoints used merely to ship faster.

## First issue set

| Issue | Owner type | Acceptance summary |
|---|---|---|
| Define Entra app registration and callback model | identity/security | ADR plus threat-model update; no code secrets |
| Implement capability/readiness evaluator | identity/platform | explicit missing consent/role/license/service/cloud states with tests |
| Build cloud-aware auth configuration | identity/platform | allowlisted authorities/hosts; cross-cloud tokens and next links rejected |
| Implement synthetic user inventory adapter | connector | pagination, schema validation, permission declaration, tests |
| Implement synthetic sign-in summary adapter | connector | time window, partial-data state, no productivity language |
| Implement synthetic usage report adapter | connector | CSV/JSON normalization and interpretation warning |
| Design tenant-scoped PostgreSQL schema | data/security | RLS policy, migrations, cross-tenant tests, deletion path |
| Run dashboard accessibility review | frontend/a11y | keyboard, focus, screen reader landmarks, contrast, motion |
| Define finding/evidence schema | domain/security | facts separated from inference; rule version and limitations |
| Create self-host deployment threat model | platform/security | secrets, updates, backup, restore, ingress, health checks |
| Implement first local AI provider adapter | AI/security | provider manifest, no action tools, bounded evidence, citations, injection tests |
