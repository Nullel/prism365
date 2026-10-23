# ADR-0006: module runtime, storage, and delegated background collection

- Status: proposed; implemented as TL-020. First adversarial review 2026-10-22: FAIL (2 critical,
  4 high, 2 medium, 2 low), all addressed below; second review pending
- Date: 2026-10-22
- Owners: founding maintainer
- Related: `docs/MODULES.md`, `docs/modules/signins-audit.md` (slice S0), ADR-0002

## Context

Prism365 needs to collect tenant data on a schedule, keep normalized history, and turn it into
findings. Every collector added ad hoc to the API would re-solve tenancy, storage, scheduling,
credentials and retention, and would widen the surface each time. The roadmap's Milestone 2
needs durable ingestion before anything can show a trend.

Two forces pull against each other. Single-operator installs (one administrator, one laptop or
small host) should not need Docker and PostgreSQL to try the product. Deployments need
PostgreSQL with a restricted runtime role and forced row-level security, which already protect
`prism365_state` and `prism365_audit`.

## Decision

1. **Modules are the only collection mechanism.** A module declares read packs, migrations in
   its own `mod_<id>_` namespace, collectors with cadence and timeout, and retention per table.
   The registry validates all of it at registration, including a static check of module SQL.
2. **One SQL dialect, two backends.** PostgreSQL when `DATABASE_URL` is configured; otherwise
   PGlite (PostgreSQL 18 compiled to WebAssembly, in process), in memory by default or on disk
   when `PRISM_DATA_DIR` is set. Both run the same migrations and the same row-level policies.
   Every `npm test` exercises the real policies on PGlite; CI exercises them on PostgreSQL 17.
3. **One tenant policy, applied by one function.** `prism365_tenant_table(regclass)` enables and
   forces row security, creates the policy on `tenant_id`, and grants the runtime role. Startup
   refuses to continue unless every table in the `prism365_` and `mod_` namespaces is a plain
   table with exactly one policy, permissive, for all commands and all roles, whose expression is
   the canonical tenant comparison, with no triggers; and unless the public schema has no
   functions other than that helper. An extra `USING (true)` policy or a definer function stops
   startup.
4. **SQL is judged on PostgreSQL's parse tree, never its text.** Every runtime statement is parsed
   with libpg-query (the PostgreSQL 17 parser) and must be one `SELECT`, `INSERT`, `UPDATE` or
   `DELETE`, reading only tables the caller is allowed and writing only tables it may write,
   calling only allowlisted functions. Module transactions may read and write their own tables
   and read findings; core transactions may use every tenant table except the session store,
   the access ledger and the migration ledger, which no module or report query can name.
   Quoting, comments and casing cannot change a parse tree, which is why the first text-pattern
   guard (bypassed with `"set_config"(...)` in review) was replaced. Module migrations are
   checked the same way: only tables, indexes, column additions and constraints in the module's
   namespace, plus the policy call.
   On PostgreSQL the restricted runtime role is a second boundary. On PGlite it is not: its
   session user stays a superuser even when it opens as another role (verified with PGlite
   0.5.11), so the parse-tree guard is the boundary there. Prism365 connects one tenant per
   installation, so a PGlite install never holds two tenants' data; multi-tenant deployments
   must use PostgreSQL.
5. **Delegated collection first.** Background collection uses the signed-in operator's
   delegated session, bound explicitly through `POST /v1/collection/start`. The session handle
   stays in process memory. Collection stops on sign-out, session expiry, removal from the
   operator allowlist, a newer sign-in, an explicit stop, or process restart. It only uses
   scopes that are both consented and enabled in `ENABLED_READ_FEATURES`. Application
   credentials (certificate, 24/7) are a later, separate identity per ADR-0002.
6. **Findings are the module output.** One shared table, deduplicated per rule and subject,
   checked against the existing finding language policy before storage. A resolved finding that
   recurs reopens; an acknowledged one reopens when its severity rises.
7. **Migrations are checksummed.** Module migrations apply once through a ledger; a migration
   whose text changed after it was applied stops startup.

## Consequences

- New capabilities arrive as reviewable folders with a fixed contract, not as API surgery.
- Single-operator installs work without Docker. In-memory mode loses baselines on every restart,
  which the server says at startup; persistent single-operator use needs `PRISM_DATA_DIR`, and
  that directory holds tenant data protected only by the host's disk encryption and file
  permissions.
- Module tables are isolated between tenants, not between modules. Modules are first-party code
  reviewed in this repository; the static SQL checks and namespace prefix prevent mistakes, not
  a malicious module.
- One scheduler per database is assumed. Abandoned-run recovery marks every `running` row for a
  tenant as abandoned on the first cycle, which would be wrong with two API instances.
- PGlite adds a WebAssembly dependency (about 10 MB) and runs PostgreSQL 18 where deployments
  run 17. SQL must stay within features both support. libpg-query adds the PostgreSQL 17 parser
  (WebAssembly, MIT); checking a statement costs about 0.15 ms.
- Transactions are bounded: a client handed to a callback stops working when its transaction
  ends, every transaction aborts after 60 seconds, and a module's transaction aborts the moment
  its run ends. A collector stuck inside one cannot hold the database or block shutdown, which
  waits at most 15 seconds for a running collector.
- A lock file in `PRISM_DATA_DIR` refuses a second process on the same data directory; two
  PGlite instances on one directory lose rows silently.
- Promises handed to collector code are marked handled, and the server logs any stray rejection
  instead of exiting, so one collector bug cannot stop collection for the whole process.
- `npm run db:migrate` now runs from `apps/api` so it can include module migrations; the old
  storage-only migrator is removed.

## Validation

- Shared isolation suite run against both backends: cross-tenant reads, writes, moves, shared
  tables, the runtime guard, and the startup verifier.
- Registry tests for namespace escapes, missing tenant policy, privilege and DDL statements.
- Scheduler tests for cadence, overlap, missing permissions, undeclared packs, transaction
  misuse, timeout with discarded late writes, credential loss, mid-run stop, notifications,
  cursor persistence, retention scoped to one tenant, and abandoned-run recovery.
- API tests for authorization, CSRF, sign-out unbinding, enabled-feature narrowing, findings
  state changes with audit entries, and demo mode.
- A mandatory adversarial review before commit.

## Revisit when

- A second API instance or a separate worker process is introduced (scheduler leasing).
- The first third-party or community module is proposed (module-level isolation, per-module roles).
- Application-credential collection lands (ADR-0002 follow-up).
- PGlite's major version or PostgreSQL compatibility changes.
