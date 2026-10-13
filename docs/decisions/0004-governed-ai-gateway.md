# ADR-0004: use a governed, provider-neutral AI gateway

- Status: accepted
- Date: 2026-10-13
- Owners: maintainers

## Context

Administrators want an AI assistant that can analyze all data available in TenantLens. Giving a
model direct database or Graph access would bypass tenant isolation, purpose limits, field
minimization, provider disclosure, and the future authorization layer. It would also turn prompt
injection in provider data into a privileged control path.

Self-hosters need freedom to choose local, tenant-hosted, or external inference without making
one vendor SDK part of the domain model.

## Decision

- Put all model calls behind a provider-neutral `AiAnalysisBroker`.
- Authorize an exact, expiring tenant/provider/purpose/dataset grant on every analysis.
- Represent “all data” as the explicit list of current dataset IDs, never a future-expanding
  wildcard.
- Retrieve bounded tenant-scoped evidence through an application-owned reader; never give the
  model Graph tokens, SQL, database credentials, or direct connector access.
- Omit direct identifiers and reject raw records unless separately authorized.
- Require structured answers with supplied evidence citations and limitations.
- Treat questions, evidence, and model responses as untrusted; enforce deterministic output
  policies outside the model.
- Provide no write tools or autonomous administrative authority.
- Record non-content access metadata in an AI audit ledger.

## Consequences

The model cannot freely explore the database and may need multiple retrieval rounds for complex
questions. Provider adapters require deliberate configuration and review. In return, the same
authorization, privacy, evidence, and tenant-isolation rules work across local and hosted models,
new datasets are not silently disclosed, and replacing a model vendor does not rewrite the
application core.
