# ADR-0005: isolate imported lifecycle automation behind capability grants

- Status: superseded for execution on 2026-10-22: Prism365 stays read-only (see `docs/ROADMAP.md`,
  Milestone 6). Private import and preview remain as design artifacts; no runner will be built.
- Date: 2026-10-13
- Owners: maintainers
- Implementation update: 2026-10-14

## Context

Organizations often have existing PowerShell onboarding and offboarding scripts. Requiring an
administrator to leave Prism365 and run them manually creates friction and inconsistent evidence.
Running an uploaded script inside the API process, or passing it an unrestricted tenant token,
would let organization code bypass Prism365 authorization, egress, tenant isolation, audit, and
recovery controls.

## Decision

- Treat scripts as immutable, versioned automation packages with a validated manifest and digest.
- Accept JSON manifest/text-file packages only, structurally inspect without extraction or
  execution, and assign `unreviewed` state on every import.
- Restrict initial imports to PowerShell 7, fixed safe paths and present `.ps1` entrypoint,
  50 files/500000 bytes, and empty direct-network host declarations.
- Hash canonical manifest plus exact source bytes; store private source encrypted and scoped
  to the authenticated tenant, with a 30-day expiration and metadata-only listing.
- Separate import, review, preview, approval, execution, and reconciliation states.
- Keep `apps/api` as the control plane; it never starts an imported process.
- Run a future package in a short-lived isolated worker with a dedicated write identity, bounded
  resources, no host access, and denied egress by default.
- Prefer a capability broker so scripts request named actions without receiving provider tokens.
- Require target resolution, protected-account checks, exact change preview, explicit confirmation,
  idempotency, append-only action records, and step-level recovery.
- Bind each approval to tenant, package digest, version, declared capability set, and parameter
  policy. Any package or capability change invalidates approval.
- Keep AI outside approval and execution authority.
- Do not expose an execution endpoint until the ledger, runner, authorization, reconciliation,
  kill switch, and independent security review are complete.

## Implemented pilot

Authenticated, CSRF-protected imports and metadata listing are integrated with encrypted
PostgreSQL state. Imported source is immutable and unreviewed; it is not added as an executable
runbook. Existing lifecycle previews remain synthetic sample plans and do not inspect tenant
targets. Review approvals, content/security scanning, signatures, isolated worker, capability
broker and all write actions remain outstanding. No ZIP extraction or process execution exists.
This initial live-auth flow requires an authorized tenant pilot before compatibility is claimed.

## Consequences

Existing scripts may need a wrapper or migration to the capability broker, and some terminal-style
behavior will not be supported. Imports take a review step before appearing in the UI. This is more
work than invoking PowerShell directly, but it prevents a convenient feature from becoming remote
code execution in the most privileged part of Prism365 and makes each change explainable and
recoverable.
