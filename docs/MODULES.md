# Modules: how Prism365 grows without becoming a monolith

**Status:** spine implemented 2026-10-22 (TL-020, ADR-0006), adversarial review pending. Supersedes
the "feature pack per endpoint" framing for anything that collects on a schedule. Permission packs
stay as they are; a module is the unit above them. Where this document and ADR-0006 differ, the
ADR records what was built.

## 1. The rule

Everything Prism365 learns about a tenant arrives through a **module**. The core never knows
what a sign-in or a mailbox rule is. It knows how to run collectors on a schedule, keep
tenant-scoped tables, store findings, render registered views, and deliver a digest. A module
brings the rest. This is what makes "all-in-one" survivable for one maintainer: the core stays
small and boring, and every capability is a folder that obeys one contract.

Three invariants, enforced by the core, not by convention:

1. **Raw tenant data stops at the ledger.** Collectors normalize before they write. Raw
   payloads are discarded unless the tenant enables a documented retention case.
2. **Modules cannot touch auth, the Graph boundary, storage internals, or the AI boundary.**
   They receive a context object and return rows, findings and view descriptors. Storage
   enforces it on the parse tree: module SQL can name only the module's own tables (and read
   findings); the session store, access ledger and migration ledger are unreachable.
3. **Findings are the only thing the AI and the digest ever read.** A finding carries its rule
   id, rule version, evidence references, time window, limitations and lifecycle state.

## 2. The contract

```ts
interface PrismModule {
  id: string;                        // "signins.audit", "mail.forwarding", "directory.changes"
  title: string;
  packs: GraphFeatureId[];           // permission packs from packages/graph/src/permissions.ts
  requires: Capability[];            // "entra.p1", "entra.p2", "role.reportsReader", "exo.appOnly"
  migrations: string[];              // SQL files; tables prefixed with the module id; core applies RLS
  collectors: Collector[];           // { id, cadence, lookback, overlap, run(ctx) }
  baselines?: BaselineUpdater[];     // learned state updated from successful observations only
  rules: Rule[];                     // pure: (observation, context) => Finding | null; versioned
  views: ViewDescriptor[];           // API routes + web panels the core mounts under the module id
  digest?: (ctx) => DigestSection;   // one section in the cross-module daily digest
  retention: { eventsDays: number; baselinesDays: number; findingsDays: number | "indefinite" };
}
```

What the core provides to every module through `ctx`:

- `graph`: the existing `GraphHttpClient`, bound to the module's declared packs. A request for
  a path outside those packs throws, exactly as today.
- `db`: a tenant-scoped transaction over the module's own tables plus the shared `finding` and
  `cursor` tables. Row-level security is applied by the core's migration helper; a module
  migration that forgets it fails CI.
- `cursor`: get and set a per-collector checkpoint (ISO timestamp or opaque string).
- `findings.upsert(finding)`: dedupes on (module, rule id, subject, evidence digest), keeps
  `first_seen` and `last_seen`, and moves state (open, acknowledged, resolved).
- `notify(finding | digest)`: fan-out to configured sinks. Sinks are operator-configured
  Teams or Slack incoming webhooks and the in-app list. No Graph write scope, ever.
- `probe`: the capability results (licence, role, consent) so a collector can skip cleanly.

Shared tables owned by the core:

- `finding` (tenant_id, module, rule_id, rule_version, severity, score, subject_kind,
  subject_id, subject_label, summary, evidence_json, window_start, window_end, limitations,
  state, first_seen, last_seen, resolved_at)
- `cursor` (tenant_id, module, collector, value, updated_at)
- `module_run` (tenant_id, module, collector, started, finished, fetched, inserted, findings,
  status, error_category)

Besides findings, a collector can report notification-center items with `context.attention.upsert`
and, after a complete read of one kind, `context.attention.retire`. The store decides what is a
meaningful change and what notifies (ADR-0007), so a connector only describes the provider's facts.

A collector's `run` returns `{ complete, more? }`. `complete: false` means records were left out
(unreadable, or too many to read). `more: true` means it stopped at its own time budget with
more to read now, as during history loading; the scheduler then runs it again within seconds
instead of waiting for its cadence.

## 3. The scheduler

One loop in the API process (later a separate worker), ticking every minute, running each
collector whose cadence is due, serially per tenant, with the module's lookback and overlap.
Credentials come from whichever of these exists, in this order:

1. **Application credential** (certificate, application permissions consented by the GA).
   Runs 24/7. This is Milestone 2's "scheduled collection uses a certificate".
2. **Operator session** (delegated, the token cache in the current session). Runs only while
   an allow-listed operator is signed in, which is enough for a pilot and for anyone who keeps
   the tab open at work. The UI says plainly "collecting while you are signed in".

Backfill is a collector run with `since = now - N days` and no overlap, flagged `backfilled`.

## 4. The catalog, in build order

| # | Module | Source | New permissions | Why this order |
|---|---|---|---|---|
| 1 | `signins.audit` | port of Entra Guard, see `modules/signins-audit.md` | none (AuditLog.Read.All already consented) | proven rules, real data at the pilot tenant today, the clearest "the portal cannot do this" |
| 2 | `mail.forwarding` | port of the forwarding monitor's Graph half: inbox rules that forward or redirect, per mailbox | application `MailboxSettings.Read` (delegated access reaches only mailboxes delegated to the admin) | waits for the application credential, see ADR-0009 |
| 3 | `directory.changes` | the directory audit log kept for good, a filterable feed, changes on each person's page, findings for risky changes | `AuditLog.Read.All` (consented) | in progress, see `docs/EYES-PLAN.md` S1 |
| 4 | `licenses` | subscribedSkus daily, reclaim candidates | `LicenseAssignment.Read.All` (consented) | cheap, immediately useful |
| 5 | `usage.reports` | the per-user Microsoft 365 apps report, which v1.0 serves as JSON, and the other per-user reports, which answer with a redirect to a CSV download that only `GraphHttpClient.getReport` follows (ADR-0010) | `Reports.Read.All` | built as an on-demand audit instead of a module, see `docs/EYES-PLAN.md` S6 |
| 6 | `service.health`, `service.messages` | built: incidents, advisories and Message center into the notification center, see `modules/service-health.md` | `ServiceHealth.Read.All`, `ServiceMessage.Read.All` | platform-level, no per-person data; first source of the notification center (ADR-0007) |

The AI analyst reads `finding` and the digest, nothing else. That is how "oversee everything
without access to the real stuff" stays true as modules are added: nothing below the ledger is
ever in its reach.

## 5. What a module may not do

- Call any host other than Graph through `ctx.graph`. Enrichment against third parties
  (reputation, ASN, geolocation) is a separate **enrichment provider** with its own consent
  toggle, egress allowlist and audit event, off by default. Tenant IP addresses leaving the
  installation is a privacy decision the operator makes once, explicitly.
- Store a raw payload unless `retention.rawEvents` is enabled for the tenant.
- Label a person. Findings describe events and risk, never productivity or intent, per
  `PRIVACY.md`. A rule named "after hours" is allowed only as "outside this account's usual
  sign-in hours", scored low, and never aggregated into a per-person score.
- Write to the tenant. There is no API for it.

## 6. Storage (decided in S0)

PostgreSQL when `DATABASE_URL` is set; otherwise PGlite, in memory by default or on disk under
`PRISM_DATA_DIR`. Same migrations, same policies; see ADR-0006 for the reasoning and the costs.

What was built differs from §2 in four places:

- Collectors get `ctx.table(name)` for their own tables and `ctx.db(fn)` for one transaction at
  a time. Graph calls, cursor and finding calls inside `ctx.db` are refused, because a held
  transaction would block the in-process database.
- `ctx.graph` exposes only `get` and `getCollection`, both requiring a response schema.
- Views and the digest are not part of the spine. The API exposes `GET /v1/modules`,
  `GET /v1/findings`, `POST /v1/findings/:id/state`, and `POST /v1/collection/start|stop`;
  module-specific views arrive with each module (S3).
- Sinks exist as an interface; no external sink ships yet (S4).
