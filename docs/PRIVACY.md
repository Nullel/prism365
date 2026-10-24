# Privacy and responsible-use principles

Prism365 handles information about people because administrators need to secure and
operate Microsoft 365. That access does not make every possible analysis appropriate.

## Core commitments

- Collect the minimum fields and time range required for an enabled feature.
- Explain collection and retention before consent.
- Keep raw records only when they are necessary and explicitly configured.
- Give tenant administrators export and deletion controls.
- Encrypt tenant-derived data in transit and at rest.
- Redact operational logs and make self-hosted product telemetry opt-in.
- Do not sell tenant data or use it to train shared models.
- Keep customer content out of fixtures, screenshots, issues, and support bundles.

## No productivity inference

Sign-ins, messages, meetings, file edits, presence, audit events, and report counts do not
measure an individual's productivity, effort, intent, quality, engagement, or performance.
They are incomplete, context-dependent operational signals. Prism365 will not rank
workers, generate productivity scores, infer “idle” employees, or present surveillance as
workforce analytics.

Usage views must:

- state the operational question they answer;
- show the aggregation and time window;
- disclose missing or delayed data;
- avoid normative labels such as lazy, inactive worker, top performer, or low engagement;
- prefer service- or tenant-level trends over individual comparisons;
- restrict person-level evidence to legitimate administration or security workflows.

## Automated explanations

An AI-generated explanation must be labelled, cite the underlying records or rules, state
uncertainty, and never invent evidence. AI may summarize or propose a query; it may not
silently expand permissions, retention, or administrative authority. Autonomous tenant
writes are out of scope.

AI access is disabled until a tenant selects an exact provider and grants named datasets. Local
or tenant-hosted inference is available as a deployment choice; external processing needs a
separate affirmative grant. Direct identifiers and raw records are separate opt-ins and remain
off by default. Adding a future dataset does not expand an existing AI grant.

Model credentials, full prompts, evidence bundles, and answers are excluded from routine logs.
The AI access ledger records non-content metadata needed for accountability. Prism365 does not
use tenant prompts, responses, or evidence to train a shared model.

## Module data

Modules declare a retention period for each table they own; the scheduler deletes older rows for
each tenant once a day. Collector run history is kept 30 days. Findings are kept until they are
deleted; a deletion route and export for findings do not exist yet and are required before a
stable release. Background collection uses the signed-in administrator's delegated access, starts
only when that administrator asks for it, and stops when they sign out.

The People view shows one account's full IP addresses, locations and devices, which the tenant
summary masks. Opening a person, paging their sign-ins, a live directory read and an email
lookup each write an access-ledger entry that names the datasets, not the person; the list of
people does not. Stored sign-in tables are encrypted at rest only by the host's disk
encryption (FileVault, BitLocker or LUKS); Prism365 does not encrypt those columns itself.
Backups made with `scripts/docker-live.sh backup` are plain `pg_dump` files outside retention.

The "Check Have I Been Pwned" button on a person page opens haveibeenpwned.com with that person's
address in the link, so the administrator's browser, not Prism365, sends it to HIBP, only when
the button is pressed. The address also lands in that browser's history.

## Audit results

Each audit run stores the rows it found in `prism365_audit_run`: account names, sign-in names,
last sign-in dates, account state, license names and, for the MFA audit, registered method types
(never the phone numbers or keys themselves). The newest 10 successful and 5 failed runs per
audit are kept and older ones deleted. Starting a run, opening a result and exporting it each write an access-ledger
entry that names the audit and the run, not an account. A downloaded CSV is outside Prism365's
retention and should be handled like the directory export it is.

## Handoffs and receipts

A handoff (`docs/handoff.md`) is a file an administrator exports from an audit for another tool to
act on: account ids, names, sign-in names, the reasons and the columns the page showed. Prism365
keeps only a record of each export (the audit, view, count, when and by whom), never its
accounts; the downloaded file is outside Prism365's retention and should be handled like the CSV
export. A receipt is that tool's report of what it did, imported as it was sent: account ids or
sign-in names, the actions, their outcomes, the operator's name as the tool gives it, and free
text. Receipts are shown as the tool's claims, never checked against the tenant.

Handoff records and receipts are kept 3 years by default, configurable with
`HANDOFF_RETENTION_DAYS` from 30 to 3650 days, measured by the database's clock. Older ones are
deleted, receipt items included, a batch at a time after each new handoff or receipt is stored;
a handoff stays while a kept receipt answers it.
Exporting, importing and reading them each write an access-ledger entry with ids and counts, not
names or the tool's text.

## Retention defaults

Final durations will be configured during the storage milestone. The design baseline is:

- normalized summaries rather than raw payloads;
- short, explicit raw-event windows when a feature genuinely requires them;
- per-category retention, not one indefinite global setting;
- deletion jobs that cover database rows, caches, exports, backups under policy, and search
  indexes;
- legal-hold support only when explicitly configured by the tenant and clearly disclosed.

## Privacy review questions

Every data feature must answer:

1. What administrator problem does it solve?
2. What exact fields are collected and under which permission?
3. Can aggregation, on-demand retrieval, or a shorter window solve the same problem?
4. Who can see it, export it, and delete it?
5. What harm could result from error, misuse, breach, or decontextualization?
6. How does the UI communicate limitations and prevent productivity inference?
