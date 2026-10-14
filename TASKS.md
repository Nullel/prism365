# Agent coordination board

This board coordinates local parallel work before GitHub issues are available. Do not add
secrets, personal data, or customer data. Once published, link each row to a GitHub issue.

| ID | State | Owner | Scope / files | Dependencies | Handoff |
|---|---|---|---|---|---|
| TL-001 | review | Codex | Identity, authenticated API integration, single-tenant setup | ADR-0002, human test tenant | Implemented 2026-10-14: MSAL/PKCE/session/CSRF, encrypted state, API integration; synthetic authorization tests pass; live tenant acceptance pending; no writes |
| TL-002 | ready | — | Synthetic user inventory adapter | TL-001 | — |
| TL-003 | ready | — | Sign-in summary adapter | TL-001 | — |
| TL-004 | ready | — | Usage reports adapter | TL-001 | — |
| TL-005 | ready | — | PostgreSQL tenant-isolation design | Threat model review | — |
| TL-006 | ready | — | Dashboard accessibility audit | Running web app | — |
| TL-007 | ready | — | Finding evidence model | Product principles | — |
| TL-008 | ready | — | Self-host deployment profile | TL-005 | — |
| TL-009 | ready | — | Capability/readiness evaluator and UI | TL-001, compatibility contract | — |
| TL-010 | ready | — | Cloud-aware identity configuration | TL-001, test tenants | — |
| TL-011 | ready | — | Service health and license inventory adapters | TL-001 | — |
| TL-012 | ready | — | Management Activity connector design | TL-005, unified-audit test tenant | — |
| TL-013 | done | Codex | Pluggable AI analysis boundary and policy tests | Privacy/threat-model review | Changed: contracts, `packages/ai`, policy, docs<br>Verified: 31 tests, checks, build<br>Impact: optional named-dataset disclosure; no Microsoft scope/write change<br>Risk: no production evidence reader/setup UI<br>Next: tenant-scoped reader and provider setup screen |
| TL-014 | done | Codex | Lifecycle automation contracts, guarded preview API/UI, and runner design | Write-safety architecture review | Changed: contracts, automation package, API/UI, ADR-0005, docs<br>Verified: 37 tests, checks, build<br>Impact: synthetic preview only; no tenant writes or new consent<br>Risk: import and execution remain unavailable<br>Next: authenticated package registry, approval ledger, and isolated worker |
| TL-015 | done | integration_review | Graph hardening and validated user investigation | TL-001 | Changed: graph, investigation contracts/tests; 28 scoped tests; no writes; next: live tenant pilot |
| TL-016 | done | ai_automation_review | Trusted grants, audited minimized AI, bounded local transport | Authenticated API | Changed: AI broker/transport/evidence/tests/docs; 37 scoped tests; no external/raw/identifier disclosure in pilot; next: configured model acceptance |
| TL-017 | done | ui_delivery_review | Tenant setup, investigation, AI and private import UI | API contracts | Changed: web components/styles/tests; 12 interface tests plus integrated browser checks; next: live role/consent/model acceptance |
| TL-018 | done | ai_automation_review | Private immutable PowerShell JSON package inspection | Authenticated API and encrypted storage | Changed: registry/contracts/API/UI; digest/path/size/isolation tests; imports unreviewed, execution disabled; next: reviewed approval and isolated runner |
| TL-019 | done | integration_review | PostgreSQL isolation and container CI verification | Disposable CI service | CI passed 2026-10-14, including all six real PostgreSQL isolation cases and container build; runtime smoke check added in publication handoff |

## Public contribution issues

- [#1: Worldwide tenant pilot](https://github.com/nlj3/prism365/issues/1) — human-authorized test tenant required; TL-001 acceptance.
- [#2: Readiness probes](https://github.com/nlj3/prism365/issues/2) — TL-009; permission changes require review.
- [#3: Persistent ingestion jobs](https://github.com/nlj3/prism365/issues/3) — next data/platform slice.
- [#4: Native lifecycle action design](https://github.com/nlj3/prism365/issues/4) — no execution until approval controls are reviewed.
- [#5: Private package review](https://github.com/nlj3/prism365/issues/5) — immutable digest and separated reviewer authority.
- [#6: Local model acceptance](https://github.com/nlj3/prism365/issues/6) — synthetic conformance first.

Claim one issue, name the files you own, and provide the handoff format below. The founding
maintainer remains the final reviewer; issues do not authorize real tenant writes or public data
disclosure.

## Claim format

Change `ready` to `claimed`, replace the owner, and add a short timestamp/branch note to the
handoff cell. Do not claim more than one substantial task at a time.

## Handoff format

```text
Changed: paths or modules
Verified: commands and manual checks
Permission/data impact: none, or exact change
Risks: known edge cases and assumptions
Next: smallest actionable follow-up
```
