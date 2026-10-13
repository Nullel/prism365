# Threat model

Status: initial design model. Update it for auth, persistence, connectors, exports, hosted
deployment, or write actions.

## Assets

- Microsoft identity tokens, client credentials, and signing/encryption keys.
- Tenant configuration, consent state, roles, and connector cursors.
- Directory, audit, usage, finding, and export data.
- AI provider credentials, tenant AI grants, evidence bundles, prompts, and generated answers.
- Administrative action intent and the future immutable action ledger.
- Software supply chain, release artifacts, and update channel.

## Trust boundaries

- Browser to Prism365 API.
- API/worker to Microsoft identity, Graph, and Purview.
- AI gateway to local, tenant-hosted, or external model providers.
- API/worker to database, queue, cache, and object storage.
- One customer tenant to another.
- Untrusted provider or issue content to logs, UI, AI, and operators.
- Contributor source to CI and released artifacts.

## Priority threats and controls

| Threat | Example | Required controls |
|---|---|---|
| Cross-tenant access | Missing scope in a query/cache | tenant key everywhere, RLS, authorization middleware, adversarial tests |
| Token theft | Token appears in logs or browser storage | server-side broker, short-lived tokens, redaction, managed secret storage, rotation |
| Consent escalation | New feature quietly adds broad scope | feature packs, permission diff, security review, separate write identity |
| Confused deputy | Low-role user triggers privileged job | local RBAC plus tenant/provider authorization, job actor context, policy decision |
| Webhook spoof/replay | Forged Graph notification | validation token flow, client-state check, provider validation, replay/idempotency control |
| Injection | Provider value rendered or used in a query | validation, contextual escaping, parameterization, no instruction following from data |
| Data exfiltration | Export URL or support bundle leaks records | scoped signed links, expiry, authorization, redaction, export ledger |
| Unsafe retry | Write repeats after timeout | no MVP writes; later idempotency, reconciliation, bounded retry, action state machine |
| Supply-chain compromise | Malicious dependency/action | lockfile, minimal dependencies, Dependabot, CodeQL, pinned reviewed actions/releases |
| Misleading findings | Missing events produce false certainty | freshness/coverage state, source/time window, deterministic rule, uncertainty label |
| Insider misuse | Valid admin uses person-level data improperly | least privilege, purpose limits, access audit, privacy UI, no worker scoring |
| AI cross-tenant context | Shared cache or retriever mixes organizations | tenant-bound grant and query, RLS, evidence tenant validation, no shared conversation state, adversarial tests |
| AI data exfiltration | External provider receives unapproved records | provider-specific grant, named datasets, external-processing opt-in, identifier/raw-record gates, egress allowlist |
| Direct/indirect prompt injection | Question or audit field asks model to reveal data or act | structured evidence boundary, no tools, bounded retrieval, output validation, injection tests, treat output as untrusted |
| Hallucinated evidence | Model invents a record or conclusion | supplied-ID citation validation, visible limitations, deterministic policy, no automatic downstream action |
| Provider retention drift | Vendor begins storing prompts differently | provider manifest, no-store where supported, periodic policy review, kill switch, tenant disclosure |

## Security assumptions

- Microsoft provider responses, identifiers, and error text are untrusted.
- Provider logs can be late, absent, filtered, licensed differently, or incomplete.
- A tenant administrator can revoke consent at any time.
- Self-hosters are responsible for infrastructure hardening, but safe defaults and a clear
  deployment guide remain project responsibilities.
- A hosted service creates materially greater legal and operational duties and needs a new
  review before launch.
- Model behavior cannot be made trustworthy by a system prompt alone. Authorization, isolation,
  output validation, and action controls live outside the model.

## Open questions before Milestone 1

- Single-tenant versus multi-tenant Entra app registration for self-hosters.
- Certificate, federated credential, or customer-provided credential options.
- Session model, local roles, break-glass access, and administrator offboarding.
- Sovereign-cloud discovery and issuer validation.
- Exact support and disclosure process before public release.
