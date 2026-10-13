# Architecture

## Design goals

TenantLens is a multi-tenant administrative system handling privileged organizational
metadata. Its architecture favors isolation, explainability, and failure containment over
connector count or ingestion speed.

## System contexts

### Control plane

The control plane manages organizations, users, roles, connector configuration, feature
flags, consent state, and job schedules. It must not return tenant data unless the caller's
organization and role are authorized.

### Data plane

The data plane retrieves Microsoft 365 data, normalizes it, calculates findings, and stores
tenant-scoped summaries. Connectors receive a short-lived credential only for the tenant
and job they are processing.

### Administration plane

Future write actions run through a separate identity and queue. Every action requires a
preview, policy decision, explicit confirmation, idempotency key, result record, and
reconciliation path. No write capability is part of the MVP.

## Target components

| Component | Responsibility | Current state |
|---|---|---|
| `apps/web` | Accessible administrator UI | Demo dashboard implemented |
| `apps/api` | Authenticated application API | Demo read API implemented |
| `packages/contracts` | Runtime-validated shared contracts | Implemented |
| `packages/graph` | Graph request, retry, paging, cloud boundary | Initial boundary implemented |
| `packages/policy` | Product and authorization guardrails | Initial finding policy implemented |
| Worker service | Scheduled ingestion and reconciliation | Milestone 2 |
| PostgreSQL | Tenant-scoped normalized data and configuration | Milestone 2 |
| Queue | Bounded jobs, retries, dead-letter handling | Milestone 2 |
| Object storage | Optional encrypted exports/raw archives | Post-MVP and opt-in |
| Action ledger | Append-only record of administrative writes | Before any write action |

Each provider adapter follows [the connector standard](CONNECTOR_STANDARD.md). Microsoft Graph,
Purview, the Management Activity API, and workload-specific administration surfaces remain
separate security boundaries even when they expose overlapping records.

## Request and ingestion flow

1. An administrator signs in through Microsoft identity.
2. The API maps the issuer and tenant claims to a TenantLens organization and local role.
3. The permission service exposes only enabled features whose required consent exists.
4. A scheduled job requests a short-lived access token from the tenant credential broker.
5. The connector calls a documented `/v1.0` endpoint through `packages/graph`.
6. The adapter validates and normalizes responses; raw payloads are discarded unless the
   tenant explicitly enabled a documented retention use case.
7. A tenant-scoped transaction upserts records and advances a delta or time cursor.
8. The finding engine stores observed facts and explainable derived findings separately.
9. The API applies organization and role filters before returning data.

## Tenant isolation

The MVP target is shared infrastructure with mandatory `tenant_id` partition keys and
database row-level security. Application queries must still include tenant scope; row-level
security is a backstop, not a replacement. Future higher-isolation deployment profiles may
use a database or installation per tenant.

Every cache key, job payload, idempotency key, export path, metric dimension, and log context
must be tenant-scoped. Tests must try cross-tenant reads, updates, pagination, background
jobs, and stale-cache access.

## Integration boundaries

- Only connector packages know remote endpoint paths and provider payload shapes.
- The domain layer consumes normalized records, never unvalidated Graph objects.
- All outbound Graph hosts are selected from a fixed cloud configuration, not user input.
- Access tokens are supplied through an interface and must never be logged or persisted in
  browser storage.
- A request declares the feature/permission pack it belongs to; permission drift is visible.
- Retries honor `Retry-After`, are bounded, and never turn an unsafe write into an
  uncontrolled replay.
- Delta queries and change notifications are used where they are supported and justified.
  Subscriptions expire and lifecycle notifications can signal reauthorization, removal, or
  missed events, so reconciliation is always required.

## Findings and evidence

A finding is not a raw event. It has:

- an observed fact and time window;
- source endpoint/query and ingestion time;
- severity and deterministic rule version;
- confidence or evidence limitations;
- affected resource references;
- recommendation and documentation link;
- lifecycle (`open`, `acknowledged`, `resolved`, `dismissed`);
- no productivity or individual-performance score.

## Deployment profiles

1. **Local demo:** synthetic data, localhost only, no tenant connection.
2. **Self-hosted single organization:** containerized API, worker, and PostgreSQL; customer
   controls keys, storage, retention, and upgrades.
3. **Hosted multi-tenant:** future profile requiring independent security review,
   operational controls, privacy terms, tenant isolation testing, and incident response.

## Observability

Use structured logs with request/job correlation IDs and provider request IDs. Do not log
tokens, request authorization headers, raw audit records, query results, UPNs, IP addresses,
or document/message content. Product telemetry must be opt-in for self-hosted installations.

## Scaling assumptions

Graph throttling varies by service and tenant. More workers do not imply more usable quota.
Schedulers must enforce per-tenant and per-connector concurrency, honor `Retry-After`, and
spread refreshes. Bulk extraction use cases should evaluate Microsoft Graph Data Connect
rather than forcing the transactional Graph API beyond its intended shape.
