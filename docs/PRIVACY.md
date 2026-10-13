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
