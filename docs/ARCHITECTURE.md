# Architecture

## Design goals

Prism365 is a multi-tenant administrative system handling privileged organizational
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
reconciliation path. Imported scripts are immutable reviewed packages and execute only in a
short-lived isolated worker, never in the API process. No write capability is part of the MVP.

## Target components

| Component | Responsibility | Current state |
|---|---|---|
| `apps/web` | Accessible administrator UI | Investigation, connection, local AI, private import, demo overview |
| `apps/api` | Authenticated application API | Single-tenant delegated pilot, CSRF/session protection, bounded read routes |
| `packages/contracts` | Runtime-validated shared contracts | Implemented |
| `packages/graph` | Validated read requests, retry/paging/cloud boundary and investigation adapter | Implemented with synthetic tests; live acceptance pending |
| `packages/policy` | Product and authorization guardrails | Initial finding policy implemented |
| `packages/ai` | Authorized minimized evidence analysis and attempt ledger | Local integration implemented; provider pilot pending |
| `packages/automation` | Runbook previews and immutable private imports | JSON import implemented; all source unreviewed; execution unavailable |
| `packages/storage` | Encrypted tenant state and restricted append-only access ledger | PostgreSQL/RLS implemented; integration verified separately in CI |
| Worker service | Scheduled ingestion and reconciliation | Milestone 2 |
| PostgreSQL report tables | Normalized historical reports and collection cursors | Milestone 2; current database stores bounded control state only |
| Queue | Bounded jobs, retries, dead-letter handling | Milestone 2 |
| Object storage | Optional encrypted exports/raw archives | Post-MVP and opt-in |
| Action ledger | Append-only record of administrative writes | Before any write action |
| Isolated automation worker | Runs approved packages with bounded capabilities | Before any imported script execution |

Each provider adapter follows [the connector standard](CONNECTOR_STANDARD.md). Microsoft Graph,
Purview, the Management Activity API, and workload-specific administration surfaces remain
separate security boundaries even when they expose overlapping records.

## Request and ingestion flow

1. An administrator signs in through Microsoft identity.
2. The API maps the issuer and tenant claims to a Prism365 organization and local role.
3. The permission service exposes only enabled features whose required consent exists.
4. A scheduled job requests a short-lived access token from the tenant credential broker.
5. The connector calls a documented `/v1.0` endpoint through `packages/graph`.
6. The adapter validates and normalizes responses; raw payloads are discarded unless the
   tenant explicitly enabled a documented retention use case.
7. A tenant-scoped transaction upserts records and advances a delta or time cursor.
8. The finding engine stores observed facts and explainable derived findings separately.
9. The API applies organization and role filters before returning data.

## AI analysis flow

1. An authorized administrator selects a provider, named datasets, lookback, record cap, and
   whether external processing, direct identifiers, or raw records are allowed.
2. The API verifies local role and an unexpired tenant/provider access grant.
3. A tenant-scoped evidence reader retrieves only the requested normalized records.
4. The broker rejects cross-tenant, out-of-window, raw, duplicate, or unauthorized evidence.
5. Direct identifiers are structurally separated and omitted unless explicitly enabled.
6. The provider receives a bounded structured evidence bundle and no Graph/database credentials
   or action tools.
7. The broker validates schema, supplied citations, and deterministic output policy before the
   answer can be displayed.
8. A non-content audit event records actor, provider, datasets, access mode, count, and time.

See [AI integration](AI_INTEGRATION.md) for the provider and privacy contract.

## Lifecycle automation flow

1. An administrator imports an immutable package; extraction and validation do not execute it.
2. Review binds the package digest and declared capabilities to allowed tenants and parameters.
3. Preview resolves the target, protected-account policy, before-state, exact change set, required
   permission packs, and recovery paths without using the write identity.
4. Explicit confirmation and any required second approval create immutable action intent.
5. A queue sends one tenant- and run-scoped capability grant to an isolated worker. The API never
   starts the process and the preferred runner never exposes a Microsoft token to the script.
6. Each step records a provider request ID and result. Unknown outcomes stop for reconciliation;
   unsafe writes are not blindly retried.
7. A read-back verifies tenant state before the run is closed or a recovery workflow is offered.

See [Lifecycle automation](AUTOMATION.md) and ADR-0005. The current UI/API end at step 3 using
synthetic metadata with execution explicitly disabled; private imports are stored as unreviewed
packages and cannot enter a live preview or run.

## Tenant isolation

The MVP target is shared infrastructure with mandatory `tenant_id` partition keys and
database row-level security. Application queries must still include tenant scope; row-level
security is a backstop, not a replacement. Future higher-isolation deployment profiles may
use a database or installation per tenant.

The first implemented installation has one configured tenant and an operator allowlist. Its
PostgreSQL control-state and audit tables force tenant row policies; the runtime refuses superuser,
RLS-bypass, owner/member-of-owner, and mutable-ledger roles. OAuth/session/cache/grant/package state
is AES-GCM encrypted and bound to tenant/kind/key. Single-use OAuth take and update-only session
refresh prevent replay and logged-out session resurrection. This is not hosted multi-tenant support.

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
- AI providers access normalized evidence only through `packages/ai`; domain/UI code does not
  send database rows or Graph responses directly to a model SDK.
- Imported runbooks are planned only through `packages/automation`; the API and AI gateway cannot
  execute packages or mint tenant write credentials for them.
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
