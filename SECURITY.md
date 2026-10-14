# Security policy

## Supported versions

Prism365 is pre-alpha. Only the latest commit on the default branch will receive fixes.
Do not connect this scaffold to a production tenant.

## Reporting a vulnerability

Do not open a public issue. Use [GitHub private vulnerability reporting](https://github.com/nlj3/prism365/security/advisories/new)
and provide only the minimum reproduction information needed. It was enabled on 2026-10-14.

Never include live access tokens, secrets, personal data, tenant IDs, directory object IDs,
or audit records. If sensitive data was exposed, revoke or rotate it before reporting.

## Maintainer response targets

- Acknowledge a valid private report within 5 business days.
- Triage severity and affected versions within 10 business days.
- Coordinate disclosure after a fix or mitigation is available.

These are goals for a volunteer project, not a service-level agreement.

## Security baseline

- No secrets in source or browser storage.
- Separate read and future write identities.
- Least-privilege, feature-scoped consent.
- Tenant isolation at every persistence and cache boundary.
- Encryption in transit and at rest for tenant-derived data.
- Redacted structured logs and bounded retention.
- Dependency review, CodeQL, and protected-branch checks.
- Explicit authorization, preview, confirmation, and immutable ledger for future writes.
