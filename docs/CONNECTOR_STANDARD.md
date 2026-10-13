# Connector standard

Every Prism365 connector must ship a versioned manifest and a narrow adapter. The manifest
is both product documentation and an enforceable contract.

## Required manifest fields

- Stable feature ID, user-facing name, source, API version, and owner.
- Supported account types and Microsoft cloud profiles.
- Exact delegated/application permissions and whether admin consent is required.
- Additional Entra/workload roles, licenses, tenant settings, and subscriptions.
- Requested fields, data classification, normalization, retention, and deletion behavior.
- Poll/delta/subscription strategy, pagination limit, retry budget, and freshness objective.
- Capability-state probes and redacted remediation messages.
- Primary Microsoft documentation URL and date last verified.
- Synthetic fixture set and opt-in live-test coverage.

## Adapter rules

1. The adapter receives a tenant-bound credential; it never chooses a tenant from request data.
2. Hosts come from the selected code-owned cloud profile. Redirects and continuation URLs are
   revalidated before use.
3. Reads are the default. A write connector requires a separate identity, policy, preview,
   approval, idempotency, ledger, and recovery design.
4. Provider payloads are untrusted input and are validated before entering the domain model.
5. Provider error bodies are not logged or returned. Keep status, a redacted request ID, and a
   classified capability state.
6. Results carry tenant ID internally, source time, collection time, coverage, freshness,
   completeness, and connector/rule version.
7. No result or aggregate may claim to measure an employee's productivity or performance.

## Review gate

A connector is not complete when the happy-path request works. It is complete when consent
revocation, missing roles/licenses, cloud incompatibility, throttling, pagination, partial
results, data deletion, and tenant isolation have passing tests and understandable UI states.
