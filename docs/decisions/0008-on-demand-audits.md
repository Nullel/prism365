# ADR-0008: on-demand audits

- Status: accepted; implemented as TL-025. Amended by TL-031: an audit can also run daily or
  weekly on the collecting administrator's sign-in while collection is on, and every successful
  run leaves a trend point that outlives the run (`docs/audits.md`, sections 11 and 12).
- Date: 2026-10-22
- Related: ADR-0006 (module runtime), ADR-0007 (notification center), `docs/audits.md`

## Context

Administrators need point-in-time answers they act on in batches: which accounts are inactive,
which licenses can be reclaimed, who lacks MFA, which guests went quiet, which products are out of
seats. These are snapshots of Microsoft's current state, not streams of events, so the module
runtime's scheduled collectors and findings are the wrong shape for them.

## Decision

1. **An audit is a definition plus a read.** It declares its table (columns, views as tag sets,
   default sort), its read packs, and a `run` that returns rows. The page, CSV export and API are
   generic over that definition, so a new audit is one file and one registry entry.
2. **Runs happen only when an administrator asks**, with that administrator's delegated session,
   through a Graph client narrowed to the audit's declared packs. No credentials, storage or
   tenant choice reach the audit.
3. **One run per audit and tenant at a time**, stored as running before it reads anything and
   closed exactly once: by its result, its failure, a 10-minute limit or shutdown. A run a stopped
   process left open is closed as interrupted on the next read.
4. **Results are kept, bounded**: the newest 10 successful and 5 failed runs per audit, in one tenant table
   (`prism365_audit_run`) with forced row-level security. Rows are stored as the audit returned
   them, validated against the contract first.
5. **Counting errs toward recent.** Where Microsoft's data is ambiguous (no recorded success),
   an account is counted from evidence that can only understate inactivity, and the basis is
   shown on every row.

## Consequences

- Audits do not raise notifications yet. A scheduled audit that feeds the notification center
  would reuse the same definitions.
- One API process per installation is assumed, as for the scheduler: a second process would
  close the first one's runs as interrupted.
- Stored results hold account names and sign-in dates for as long as they are among the newest
  10 runs.
