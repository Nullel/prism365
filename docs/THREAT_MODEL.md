# Threat model

Status: initial design model. Update it for auth, persistence, connectors, exports, hosted
deployment, or write actions.

## Assets

- Microsoft identity tokens, client credentials, and signing/encryption keys.
- Tenant configuration, consent state, roles, and connector cursors.
- Directory, audit, usage, finding, and export data.
- AI provider credentials, tenant AI grants, evidence bundles, prompts, and generated answers.
- Script packages imported before the import was removed (2026-10-24), encrypted until they
  expire or an operator deletes them, and backups that still hold them.
- The mailbox access export script (`/scripts/Export-MailboxAccess.ps1`), which Prism365 serves and
  administrators run with their own Exchange Online session.
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

Prism365 has no write path (ROADMAP Milestone 6). The threats that come with acting on a tenant,
such as a malicious imported script, a runner escape, targeting a protected account or a lifecycle
run that fails halfway, belong to the threat model of the action tool that a handoff feeds.

### Imported mailbox access

Shared mailbox access comes from a file an administrator exports and imports (ADR-0012), so the
file is untrusted input:

- An edited or forged file can show access that does not exist, or hide access that does. Prism365
  cannot check it against Exchange. Every page shows where the data came from and how old it is,
  warns after 30 days, and refuses exports older than 90 days, dated in the future, or from
  another tenant. Nothing in Prism365 acts on the import or feeds it to audits, alerts or
  handoffs, so a false file misleads a reader but triggers nothing.
- A hostile file cannot reach storage or the page as anything but bounded text: the whole file is
  validated before anything is recorded (strict format, GUID ids, every string bounded on the text
  as sent and cleaned of controls and direction marks, at most 20,000 mailboxes and 200,000 grants,
  25 MB on that route only, read only for a signed-in administrator), queries are parameterised,
  and the page renders it as text, never HTML.
- A replace that fails partway keeps the previous import whole (one transaction), two imports
  cannot interleave (a row lock and a one-import-per-tenant index), and tables are tenant-scoped
  with row-level security like every other.
- The script runs with the administrator's Exchange session. It calls only Get- cmdlets, it never
  looks a trustee up with an empty identity (which Exchange answers with every recipient in the
  tenant), and it stops, writing nothing, on any Exchange error other than a trustee that is not
  found or matches more than one recipient. Tests read it with PowerShell's own parser and fail on
  any command not on a short list, any call or dot-source operator, any redirection and any module
  import; a text check that does not need PowerShell refuses the same classes of syntax.
- Prism365 serves code that administrators run with Exchange admin rights, so a changed copy is a
  way into Exchange: an attacker who can change the image, the files under `WEB_DIST_PATH`, or a
  proxy in front of the server could serve a script that writes. A changed script can do anything
  the administrator's session can; Prism365's own read-only guarantees do not cover it. Controls:
  the script's SHA-256 is published in `docs/SETUP.md` (a test fails when the file and the
  published hash differ), the page works out the hash of the copy it serves in the browser and
  warns when it is not the one this version shipped, and both the page and the guide tell the
  administrator to read the script and compare `Get-FileHash` against the hash on GitHub before
  running it. The page's own check cannot be trusted if the server is the one that was changed,
  which is why the comparison with GitHub is the one the guide relies on.

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
- Organization-owned scripts are untrusted executable code even when their publisher is known.
  Manifest declarations limit approval; they do not replace isolation or runtime enforcement.

## Open questions before Milestone 1

- Single-tenant versus multi-tenant Entra app registration for self-hosters.
- Certificate, federated credential, or customer-provided credential options.
- Session model, local roles, break-glass access, and administrator offboarding.
- Sovereign-cloud discovery and issuer validation.
- Exact support and disclosure process before public release.

## Pilot controls implemented

The single-tenant pilot implements confidential MSAL authorization code/PKCE; browser-bound one-use
state; signed ID-token issuer/audience/tenant/nonce validation; an explicit operator object-ID
allowlist; opaque HttpOnly/Lax cookies with Secure on HTTPS; Origin and CSRF checks for mutations;
login and authenticated request limits; and server-only token caches.

The Graph reader rejects writes/redirects, enforces named feature/endpoint boundaries, validates
provider schemas, caps time/response/paging, and defers rather than shortening provider throttles.
Its one redirect is the usage report download (ADR-0010): only to `reports.office.com` under
`/data/download/` over HTTPS, without the token, cookies or further redirects, used once, capped in
bytes and parsed by a bounded CSV reader.
The AI broker loads trusted expiring grants, checks current actor permissions, records attempts
before evidence access/disclosure, minimizes fields, preserves structural coverage, and rejects
invented citations and prohibited workforce claims. Its local transport rejects redirects and caps
streamed responses.

Alert delivery (ADR-0011) is the one path that sends data out on the server's own initiative. It
never uses Graph. Webhook addresses are HTTPS host names only, resolved at send time with every
answer checked against loopback, private, link-local, CGNAT, multicast, reserved and metadata
ranges in IPv4 and IPv6, then connected to by the checked address, so DNS rebinding cannot move
the request; no redirects, a 10 second deadline and a 64 KB read cap. Addresses and signing
secrets are encrypted and bound to their row, requests are HMAC-signed, SMTP requires TLS and
reads its password only from a file, recipients are limited to `ALERT_EMAIL_DOMAINS`, and full IP
addresses leave only for email with details on. The sink only buffers, so a stuck destination
cannot stall collection.

PostgreSQL enforces tenant policies and restricts runtime ledger updates/deletes. State encryption
binds tenant/kind/key; atomic consumption and update-only refresh block OAuth replay and logged-out
session resurrection. Expired state is pruned; audit retention, key rotation, restore testing,
exhaustive readiness probes, independent security review and live tenant/model acceptance remain
release requirements. These controls are implementation evidence, not a guarantee against all
deployment or model threats.
