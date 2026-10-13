# Human and AI agent workflow

## Durable work units

A GitHub issue is the unit of intent. It describes one outcome, acceptance checks, data and
permission impact, non-goals, and dependencies. A pull request is the unit of review. An ADR
is the unit of durable architectural decision.

Before GitHub publishing, `TASKS.md` mirrors only enough state to avoid local collisions.

## Parallel contribution model

Prefer boundaries that minimize shared-file edits:

- connector agent: one normalized adapter plus tests;
- frontend agent: one view or component plus accessibility checks;
- policy agent: one rule or authorization decision plus negative tests;
- documentation agent: one reviewed specification or runbook;
- reviewer agent: read-only review with file/line findings, not silent rewrites.

Changes to root configuration, shared contracts, permissions, authentication, tenancy, and
storage are integration work. They need a named human owner and should not be split across
uncoordinated agents.

## Prompt packet for an agent

```text
Issue/outcome:
Allowed scope and owned files:
Non-goals:
Acceptance checks:
Permission and data classification:
Required documents/ADRs:
How to verify:
Expected handoff format:
```

## Review expectations

Review generated code as untrusted contribution. Verify APIs against primary documentation,
run tests, inspect dependency and permission changes, and check that error paths do not leak
tenant data. A fluent explanation is not evidence that code is safe.

## Handoff

An agent handoff names changed files, verification, permission/data impact, unresolved risks,
and one smallest next step. Never report a deployment, public repository, tenant connection,
or passed check unless an observed tool result confirms it.
