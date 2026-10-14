# Connect a non-production Microsoft 365 tenant

Prism365 now implements single-tenant delegated sign-in and a bounded user investigation. It
remains pre-alpha: automated fixtures verify the boundaries, but a live Microsoft tenant pilot
is required before claiming compatibility. Start with a non-production Worldwide work or school
tenant. Personal Microsoft accounts and sovereign-cloud login are not supported in this slice.

## Local demonstration

Install Node.js 24+ and npm 11+, then run:

```bash
npm ci
npm run dev:api
npm run dev:web
```

Open `http://localhost:5173`, select User Intelligence, and investigate the synthetic account.
All demonstration records remain labelled. An unavailable API is shown as an error; live mode
does not substitute synthetic tenant records.

## App registration and operator authorization

Create your own single-tenant Entra app registration. Configure its platform as **Web**, because
Prism365 redeems authorization codes on the server. The local redirect is exactly:

```text
http://localhost:5173/api/v1/auth/callback
```

For deployment it is `https://YOUR-PRISM365-HOST/api/v1/auth/callback`. Set `PUBLIC_ORIGIN` to
that exact origin. Do not enable implicit token grants or register the callback as a SPA.

Copy `.env.example` to an ignored `.env` and supply tenant/client IDs. Set
`ENTRA_ADMIN_OBJECT_IDS` to the comma-separated object IDs of the specific operators allowed
to use this installation. A valid account in the tenant alone does not grant Prism365 access.
The local operator allowlist never replaces Microsoft's own role and license checks.

For a local development pilot, a development client secret is supported. For deployment,
upload your public certificate to the Entra registration and configure the matching PEM
certificate/private-key paths on the server. Secret files must be outside Git and readable only
by the service account. Do not paste credentials into the UI, public issues, or this repository.

Set `DATA_MODE=live`. Enable only the selected packs in `ENABLED_READ_FEATURES`:

| Feature ID | Delegated permission | What it enables |
|---|---|---|
| `directory.users.read` | `User.Read.All` | Selected user profile fields |
| `user.licenses.read` | `LicenseAssignment.Read.All` | Per-user license details, subject to a supported Entra role |
| `audit.signins.read` | `AuditLog.Read.All` | Bounded recent sign-in evidence, subject to role/license/retention |

`User.Read` is also requested for interactive sign-in. No application or write permissions are
requested by this pilot. Sign in from the connection card; each optional pack has its own consent
link. Existing enabled grants are retained when requesting another selected pack. A provider 403
does not conclusively identify which prerequisite failed; the UI explains that uncertainty.

The session uses PKCE, a one-use browser-bound state, nonce, signature/issuer/audience/tenant
verification, and a signed operator identity. The browser receives only an opaque HttpOnly
session handle and a CSRF value. Microsoft tokens stay in the server's MSAL cache. Logout removes
that cache from Prism365; it does not revoke every Microsoft session or the app's tenant consent.
Remove consent in Entra when disconnecting the registration itself.

## Encrypted durable state

Local development may use memory sessions, which reset on restart. AI and script import require
PostgreSQL, and production live mode additionally requires encrypted durable state and certificates.

Provision PostgreSQL 17 with a separate migration administrator. Generate unique URL-safe database
passwords and a random 32-byte hex `STATE_ENCRYPTION_KEY` through your approved secret-management
process. Configure `DATABASE_ADMIN_URL`, `DATABASE_APP_PASSWORD`, and run `npm run db:migrate`.
Then set `DATABASE_URL` to the restricted `prism365_app` login. The runtime must not receive the
database administrator URL/password. Migration does not rotate an existing runtime password.

The application role must not be a superuser, bypass RLS, or own the tables. State rows use tenant
keys and forced row-level policies; encrypted values are bound to tenant/kind/key with AES-GCM.
Token-cache entries expire with their eight-hour session; OAuth state lasts five minutes, AI
grants one hour, and imported script packages thirty days. Expired state is physically pruned
on startup and every minute while the service runs. Stopped installations require maintenance
before reuse. The runtime can insert/read audit records but cannot update/delete them.

Keep the encryption key backed up securely with the encrypted database backup. Losing the key
makes stored sessions and script packages unreadable. Key rotation, backup/restore, and audit-ledger
retention/deletion procedures still require an operator-reviewed deployment runbook before stable
production use.

## Self-host container

For the synthetic demo:

```bash
docker compose up --build -d app
```

Open `http://localhost:8080`. For live mode, configure HTTPS ingress and `PUBLIC_ORIGIN`, your
certificate files under ignored `secrets/`, database passwords, and the encryption key. Start
the optional database and run the one-time migration before restarting the application:

```bash
docker compose --profile live up -d --wait db
docker compose --profile live run --rm migrate
docker compose up --build -d app
```

The application port binds to loopback. Terminate HTTPS in a trusted reverse proxy on the same
host. The application runs as a non-root user with a read-only filesystem, removed capabilities,
and no Docker socket. Its server serves both the built UI and `/api/v1` routes.

The base container images follow their Node/PostgreSQL major tags. Before a supported release,
pin reviewed image digests and test upgrades. Container building is checked in CI; no deployment
is performed by CI.

## Optional local AI

Configure `AI_LOCAL_MODEL` to an already installed Ollama model. The API calls only
`http://127.0.0.1:11434/api/chat`; redirects are rejected. When containerized, the optional
`ai` Compose profile shares the application network namespace so this endpoint remains loopback:

```bash
docker compose --profile ai up -d ollama
```

Install your chosen model through that service using its documented administration process.
The example Ollama image uses a moving tag for experimentation; pin a reviewed digest for a
deployment. Model downloads and inference hardware are operator responsibilities.

After investigating a user, select the available datasets and approve a one-hour grant in the
AI panel. The first integration permits directory, license-count, and sign-in evidence only;
external processing, raw records, and direct identifiers remain off. Arbitrary names typed in a
question cannot be reliably detected by redaction; do not put unnecessary sensitive text there.
AI can explain evidence and cannot change the tenant or run imported scripts.

## Private script import

The Lifecycle Automation view imports organization-owned JSON packages into encrypted storage.
Use `{ "manifest": {...}, "files": [{ "path": "scripts/run.ps1", "content": "..." }] }`.
Only PowerShell 7 packages with a fixed `.ps1` entrypoint, safe unique paths, empty network-host
list, and a total canonical payload below 500000 bytes are accepted. See
[Lifecycle automation](AUTOMATION.md) for the full schema and a valid example.

Each upload receives a digest and unreviewed status. Neither import nor viewing the library
executes code. Package source is excluded from list responses, AI evidence, and public repository
content. Import only code you have permission to use and remove embedded secrets first.

## Required pilot acceptance

Before enabling real organizational use, verify in a non-production tenant:

1. Allowed operator signs in; wrong tenant/operator/state/nonce is rejected.
2. User lookup succeeds; licenses/sign-ins independently explain missing access.
3. Pagination caps, empty results, throttling, and provider failures show honest coverage.
4. Removing an operator, signing out, or revoking consent blocks further access.
5. Sessions survive a server restart only when durable state is configured; no token is logged.
6. AI receives only selected minimized evidence and records failed attempts as well as success.
7. Package source remains private and cannot execute.

## Primary references

- [MSAL Node token acquisition](https://learn.microsoft.com/en-us/entra/msal/javascript/node/acquire-token-requests)
- [Microsoft authorization code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Get user](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0)
- [List user license details](https://learn.microsoft.com/en-us/graph/api/user-list-licensedetails?view=graph-rest-1.0)
- [List sign-ins](https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0)
