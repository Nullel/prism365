# Set up Prism365 for your tenant

Prism365 is a self-hosted, read-only console for one Microsoft 365 tenant: sign-in monitoring,
one page per account, on-demand audits, and Microsoft's service health and Message center. It
signs in with your organization's own Entra app registration and only ever reads. It is
pre-alpha, so start with a non-production Worldwide work or school tenant. Personal Microsoft
accounts and sovereign clouds are not supported.

Getting from a clone to a working install is one command and one consent click. The command
registers the app, makes its certificate and writes `.env`; the click, granting the
permissions, stays yours on purpose.

## Quick start

You need Node.js 24 or newer, npm 11 or newer, Docker, and an account that can create app
registrations in the tenant (Application Administrator, Cloud Application Administrator or
Global Administrator).

1. Install and run the setup with your tenant's ID or one of its domains:

   ```bash
   npm ci
   npm run setup -- --tenant contoso.onmicrosoft.com
   ```

   It prints a device code and a sign-in address; sign in there with the administrator
   account. It signs in through Microsoft's own public client for the Microsoft Graph command
   line tools and asks for `Application.ReadWrite.All` and `User.Read`, so the first time
   Microsoft may ask you to consent to that tool. The token stays in memory and is never
   written to disk.

2. Read the plan it prints, and type `yes` to apply it. It then:

   - registers a single-tenant web app (tagged `prism365-setup`, implicit grant off) with the
     redirect URI `https://localhost:8443/api/v1/auth/callback` and the delegated Microsoft
     Graph permissions the chosen features need, and creates its service principal;
   - makes an RSA 2048 key and a self-signed certificate valid for 365 days, writes
     `secrets/entra-private.pem` (readable only by you, in a folder only you can open) and
     `secrets/entra-public.pem`, and adds the certificate to the registration, keeping every
     certificate already there;
   - creates `.env` from `.env.example`, or changes only the settings it manages in an existing
     one: `DATA_MODE`, `ENTRA_TENANT_ID`, `ENTRA_CLIENT_ID`, `ENTRA_ADMIN_OBJECT_IDS` (your
     account is added, others stay), `ENABLED_READ_FEATURES`, `PUBLIC_ORIGIN`,
     `ENTRA_CERTIFICATE_PATH` and `ENTRA_PRIVATE_KEY_PATH`.

3. Open the admin consent link it prints, sign in as an administrator, and accept. Microsoft
   then sends the browser to Prism365's address, which may not be running yet; if that page
   does not load, the consent is still saved.

4. Start the live install:

   ```bash
   scripts/docker-live.sh init
   scripts/docker-live.sh up
   ```

5. Open `https://localhost:8443`, sign in, and look at the setup checklist on the Microsoft
   access page.

The setup never grants consent, never creates a client secret, never asks for an application
permission, and never removes a permission or a certificate it did not add, unless you ask with
`--prune-permissions`. It is the only part of the project that changes anything in Microsoft;
Prism365 itself only reads.

### Setup options

Pass them after `--`, as in `npm run setup -- --tenant contoso.onmicrosoft.com --dry-run`.

| Option | What it does |
|---|---|
| `--tenant <id or domain>` | The tenant to set up. Required. |
| `--name <display name>` | The registration's name. Default `Prism365`. |
| `--origin <origin>` | Where Prism365 runs. Default `https://localhost:8443`, the Docker install. The redirect URI is this origin plus `/api/v1/auth/callback`, and it becomes `PUBLIC_ORIGIN`. |
| `--dev` | Also registers the development server's redirect URI, `http://localhost:5173/api/v1/auth/callback`. |
| `--features <list>` | Comma-separated read packs to enable. Default: the packs the audits and modules use, plus the ones the People page reads. See the table below. |
| `--app-id <application ID>` | The registration to use when several match. |
| `--secrets-dir <dir>` | Where the certificate goes. Default `./secrets`. |
| `--env-file <file>` | The settings file to write. Default `.env`. |
| `--rotate-certificate` | Makes a new key and certificate even when a pair exists. |
| `--prune-permissions` | Removes permissions on the registration that the chosen features do not use. |
| `--dry-run` | Signs in, reads, prints the plan and stops. Nothing is written anywhere. |
| `--yes` | Applies the plan without asking. |

### Running it again

Every step checks what is already there first, so running the setup again is safe and finishes
a run that stopped halfway. It finds its registration by the `prism365-setup` tag, then by its
exact name; if several match, it stops and asks for `--app-id`. On an existing registration it
adds missing redirect URIs, permissions and the tag, and leaves everything else as it is. A
second run on a finished install changes nothing.

The certificate's private key is never overwritten without `--rotate-certificate`. If the key
exists without its certificate, the setup makes a certificate for that same key. If the two
files do not belong together, or the certificate has expired, it stops and says so.

Adding a certificate to a registration replaces the registration's whole certificate list, so
the setup reads the list with every key, sends each one back unchanged with the new one, and
checks afterwards that every certificate from before is still there. If Microsoft does not
return a key in full, it stops without changing anything.

### Rotating the certificate

The certificate lasts 365 days, the longest Microsoft accepts through its API. The setup
checklist warns 30 days before it expires. To replace it:

```bash
npm run setup -- --tenant contoso.onmicrosoft.com --rotate-certificate
scripts/docker-live.sh down
scripts/docker-live.sh up
```

The old certificate stays on the registration until you remove it under Certificates and
secrets in the Entra admin center.

## The setup checklist

The Microsoft access page shows what this installation has set up and what it has not, each
with its fix. Without a signed-in session it says only what the page already shows: whether the
installation is configured, and which settings are missing. Signed in, it checks:

- the configuration is complete;
- the credential is a certificate (a client secret warns: it is for development only);
- the certificate's expiry, read from the public certificate when Prism365 starts (it warns
  under 30 days and fails once expired; the private key is never read for this);
- you are signed in and on `ENTRA_ADMIN_OBJECT_IDS`;
- each enabled read pack has been granted by your session, with its consent link when not;
- sessions and collected data survive restarts;
- collection is running, or will resume;
- each audit has run at least once.

Check with Microsoft makes one small read per granted pack (one item, and only its ID where
the endpoint allows) and shows Microsoft's answer. It writes a `setup.probed` entry naming the
packs to the access ledger before it reads. When Microsoft refuses a read even though the
permission is granted, the check names the roles that pack needs.

## Manual setup

Do this instead of `npm run setup` if you prefer to make each change yourself.

### 1. Register the app

In the [Microsoft Entra admin center](https://entra.microsoft.com), open App registrations and
choose New registration:

1. Name it, for example `Prism365`.
2. Supported account types: accounts in this organizational directory only (single tenant).
3. Redirect URI: platform Web, `https://localhost:8443/api/v1/auth/callback` for the Docker
   install, or `https://YOUR-PRISM365-HOST/api/v1/auth/callback` for your own host. Add
   `http://localhost:5173/api/v1/auth/callback` too if you will run the development servers.
   Prism365 redeems codes on the server, so do not register it as a single-page application,
   and leave implicit grant off.

Note the Application (client) ID and the Directory (tenant) ID on its overview page.

### 2. Add the permissions

Under API permissions, add Microsoft Graph delegated permissions: `openid`, `profile`,
`offline_access` and `User.Read` for sign-in, plus the permission of each read pack you will
enable (see [What each feature needs](#what-each-feature-needs)). Do not add application
permissions. Then choose Grant admin consent, or use the consent link
`https://login.microsoftonline.com/<tenant ID>/adminconsent?client_id=<client ID>`.

### 3. Make the certificate

Create a folder only you can open, a 2048-bit RSA key and a self-signed certificate valid for
365 days. The key must be a PEM file that starts with `-----BEGIN PRIVATE KEY-----`.

On macOS and Linux, with OpenSSL (or the LibreSSL that ships with macOS):

```bash
mkdir -p secrets && chmod 700 secrets
(umask 077 && openssl req -x509 -newkey rsa:2048 -sha256 -days 365 -nodes \
  -subj "/CN=Prism365 contoso.onmicrosoft.com" \
  -keyout secrets/entra-private.pem -out secrets/entra-public.pem)
chmod 644 secrets/entra-public.pem
head -1 secrets/entra-private.pem
```

If the last line prints `-----BEGIN RSA PRIVATE KEY-----`, convert the key:

```bash
(umask 077 && openssl pkcs8 -topk8 -nocrypt -in secrets/entra-private.pem -out secrets/key.tmp \
  && mv secrets/key.tmp secrets/entra-private.pem)
```

On Windows, with PowerShell 7.4 or newer (Windows PowerShell 5.1 cannot export PEM files):

```powershell
$cert = New-SelfSignedCertificate -Subject "CN=Prism365 contoso.onmicrosoft.com" `
  -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 2048 `
  -HashAlgorithm SHA256 -KeyExportPolicy Exportable -NotAfter (Get-Date).AddDays(365)
New-Item -ItemType Directory -Force secrets | Out-Null
$key = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
Set-Content -Path secrets\entra-private.pem -Value $key.ExportPkcs8PrivateKeyPem() -NoNewline
Set-Content -Path secrets\entra-public.pem -Value $cert.ExportCertificatePem() -NoNewline
icacls secrets\entra-private.pem /inheritance:r /grant:r "$($env:USERNAME):F"
Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)"
```

The last line removes the certificate from your Windows certificate store, so the PEM files are
the only copies. If `ExportPkcs8PrivateKeyPem` is refused, use the OpenSSL commands above;
Git for Windows includes OpenSSL.

Then, on the registration, open Certificates and secrets, choose Certificates, and upload
`secrets/entra-public.pem`. Never upload or share `entra-private.pem`.

### 4. Write .env

Copy `.env.example` to `.env` (it is git-ignored) and set:

```text
DATA_MODE=live
PUBLIC_ORIGIN=https://localhost:8443
ENTRA_TENANT_ID=<Directory (tenant) ID>
ENTRA_CLIENT_ID=<Application (client) ID>
ENTRA_ADMIN_OBJECT_IDS=<your account's object ID>
ENABLED_READ_FEATURES=directory.users.read,user.licenses.read,audit.signins.read
ENTRA_CERTIFICATE_PATH=/full/path/to/secrets/entra-public.pem
ENTRA_PRIVATE_KEY_PATH=/full/path/to/secrets/entra-private.pem
```

`ENTRA_ADMIN_OBJECT_IDS` lists, comma-separated, the object IDs of the administrators allowed to
use this installation; your own is on your profile in the Entra admin center. A valid account
in the tenant alone does not get in, and the list never replaces Microsoft's own role and
licence checks. Use full paths for the certificate: the development server runs from
`apps/api`. Keep `chmod 600` on `.env`.

### 5. Start it

Run `scripts/docker-live.sh init` and `scripts/docker-live.sh up` (see
[Live install in Docker](#live-install-in-docker)), or the development servers (see
[Developing against a tenant](#developing-against-a-tenant)).

## What each feature needs

Each read pack in `ENABLED_READ_FEATURES` needs its delegated Microsoft Graph permission with
admin consent, and the signed-in administrator needs a role Microsoft accepts for it. Several
packs share a permission; Prism365 still only uses a pack you enabled (see
[Permissions](PERMISSIONS.md)). The source of truth is `packages/graph/src/permissions.ts`, and
a test keeps this table in step with it.

| Feature | Delegated permission | Role Microsoft also requires | Licence and limits |
|---|---|---|---|
| `directory.users.read` | `User.Read.All` | Limited by the signed-in administrator's own directory rights | Endpoint-specific Entra licensing applies |
| `user.licenses.read` | `LicenseAssignment.Read.All` | Directory Readers, Guest Inviter, License Administrator, User Administrator or a suitable custom role | Shows direct and group assignments |
| `tenant.licenses.read` | `LicenseAssignment.Read.All` | A directory role such as Directory Readers | Lists purchased products and seats |
| `audit.signins.read` | `AuditLog.Read.All` | A supported Entra role such as Reports Reader, Security Reader or Global Reader | Fields and retention depend on Entra ID licensing |
| `reports.authentication-methods.read` | `AuditLog.Read.All` | Reports Reader, Security Reader, Security Administrator or Global Reader | Leaves out disabled accounts |
| `service.health.read` | `ServiceHealth.Read.All` | Global Reader or Service Support Administrator | Only services the tenant subscribes to |
| `service.messages.read` | `ServiceMessage.Read.All` | Message Center Reader or Global Reader | Only services the tenant subscribes to |
| `audit.directory.read` | `AuditLog.Read.All` | Reports Reader, Security Reader or Global Reader | Kept 7 days without Entra ID P1 or P2, 30 days with it |
| `identity.conditional-access.read` | `Policy.Read.All` | Security Reader or Global Reader | Some outcomes need Entra ID P1 or P2 |
| `identity.risk.read` | `IdentityRiskyUser.Read.All` | A supported Entra role | Needs Entra ID P2 |
| `devices.intune.read` | `DeviceManagementManagedDevices.Read.All` | An Intune role | Needs an active Intune licence |
| `directory.applications.read` | `Application.Read.All` | Global Reader or Cloud Application Administrator | No extra licence |
| `directory.permission-grants.read` | `Directory.Read.All` | Global Reader | Reads the whole directory; enable it only to see delegated consents |
| `reports.usage.read` | `Reports.Read.All` | A reports or directory role, such as Reports Reader | Depends on the tenant's subscriptions |
| `reports.settings.read` | `ReportSettings.Read.All` | A supported role | Reads the name-concealment setting only |

A provider 403 does not always say which prerequisite failed; the checklist's Check with
Microsoft shows Microsoft's own answer for each granted pack.

## Local demonstration

```bash
npm ci
npm run dev:api
npm run dev:web
```

Open `http://localhost:5173`. Demo mode shows labelled synthetic data and needs no tenant. An
unavailable API is shown as an error; live mode never substitutes synthetic records.

## Developing against a tenant

The development servers serve the app at `http://localhost:5173`, so run the setup with that
origin, or with `--dev` to register both addresses and then set
`PUBLIC_ORIGIN=http://localhost:5173` in `.env` while you develop:

```bash
npm run setup -- --tenant contoso.onmicrosoft.com --origin http://localhost:5173
npm run dev:api
npm run dev:web
```

The certificate works in development too. A client secret (`ENTRA_CLIENT_SECRET`) is accepted
for local development only; production mode refuses it, and the setup never makes one.

## Sign-in and sessions

The session uses PKCE, a one-use browser-bound state, nonce, signature, issuer, audience and
tenant verification, and a signed operator identity. The browser receives only an opaque
HttpOnly session handle and a CSRF value. Microsoft tokens stay in the server's MSAL cache.
Logout removes that cache from Prism365; it does not revoke every Microsoft session or the app's
tenant consent. Remove consent in Entra when disconnecting the registration itself.

Each optional pack also has its own consent link on the Microsoft access page; existing grants
are kept when you grant another.

## Encrypted durable state

Local development may use memory sessions, which reset on restart. AI and script import require
PostgreSQL, and production live mode additionally requires encrypted durable state and
certificates.

Provision PostgreSQL 17 with a separate migration administrator. Generate unique URL-safe
database passwords and a random 32-byte hex `STATE_ENCRYPTION_KEY` through your approved
secret-management process. Configure `DATABASE_ADMIN_URL`, `DATABASE_APP_PASSWORD`, and run
`npm run db:migrate`. Then set `DATABASE_URL` to the restricted `prism365_app` login. The
runtime must not receive the database administrator URL or password. Migration does not rotate
an existing runtime password.

The application role must not be a superuser, bypass RLS, or own the tables. State rows use
tenant keys and forced row-level policies; encrypted values are bound to tenant, kind and key
with AES-GCM. Token-cache entries expire with their eight-hour session; OAuth state lasts five
minutes, AI grants one hour, and imported script packages thirty days. Expired state is
physically pruned on startup and every minute while the service runs. The runtime can insert and
read audit records but cannot update or delete them.

Keep the encryption key backed up securely with the encrypted database backup. Losing the key
makes stored sessions and script packages unreadable. Key rotation, backup and restore, and
audit-ledger retention and deletion procedures still need an operator-reviewed runbook before
stable production use.

## Module collection

Live mode also starts the module runtime (ADR-0006). It stores module data, findings, collector
cursors and run history in tenant-scoped tables with forced row-level security:

- With `DATABASE_URL` configured, it uses that PostgreSQL database. Run `npm run db:migrate`
  after upgrades; it applies the core schema and every built-in module's migrations, and the
  runtime refuses to start if a module's tables are missing.
- Without it, it uses PGlite, an in-process PostgreSQL. Data is in memory and lost on restart
  unless `PRISM_DATA_DIR` names a directory. That directory then contains tenant data; keep it
  on an encrypted disk, outside cloud-synced folders, readable only by the service account. A
  relative path resolves from `apps/api` under `npm run dev:api`, so prefer an absolute one.

Background collection never starts on its own. An allow-listed administrator starts it with
Start collecting on the Sign-in audit page (`POST /api/v1/collection/start`) and stops it with
Stop collecting. Starting is a standing choice: collection runs on that administrator's
delegated session, and when it pauses (a server restart, a sign-out, an expired session) it
resumes on their next signed-in visit. It stops for good only when an administrator presses Stop
collecting, and it
never runs for an administrator who was removed from `ENTRA_ADMIN_OBJECT_IDS`. It uses only
packs that are both enabled in `ENABLED_READ_FEATURES` and granted. `GET /api/v1/modules` shows
each module's missing permissions and its collectors' last runs. The first built-in module is
the sign-in audit (`docs/modules/signins-audit.md`); the People view builds on what it stores.

## Service health, Message center and notifications

The Overview's "Is Microsoft down?" panel, Admin attention and the Notifications inbox read
Microsoft's service communications. Each source is optional: enable `service.health.read`
and `service.messages.read` (both in the default feature list of `npm run setup`), grant them,
and start collecting. Your account needs a role Microsoft accepts, such as Global Reader,
Service Support Administrator or Message Center Reader.

Notification settings (Notifications, then Settings) choose which kinds of item notify and which
services count as important. See `docs/modules/service-health.md` and ADR-0007.

## Audits

Audits, in the sidebar under Audit, runs point-in-time checks when you press Run: inactive
accounts by days since last sign-in, licenses to reclaim, MFA registration, guest accounts and
license capacity. They use `directory.users.read`, `audit.signins.read`,
`tenant.licenses.read` and `reports.authentication-methods.read`, all in the setup's default
list. Inactivity needs Microsoft Entra ID P1 or P2 and a role such as Reports Reader, Security
Reader or Global Reader. Results are kept (the last 10 successful runs per audit) and can be
exported as CSV. See `docs/audits.md`.

## Breach lookup (Have I Been Pwned)

Each person page has a Check Have I Been Pwned button. It opens haveibeenpwned.com in a new tab
with the person's mailbox address (or, failing that, their sign-in name) already searched,
using HIBP's documented link `https://haveibeenpwned.com/account/{address}`. Nothing needs
configuring and no API key is used: the search runs in your own browser, and Prism365 itself
sends nothing to HIBP. HIBP's public search does not show breaches it classes as sensitive.

## Live install in Docker

`scripts/docker-live.sh` runs the live app with everything it collects kept in PostgreSQL:
sessions (encrypted), collected sign-ins, baselines and findings survive restarts, image
rebuilds and `down`. It layers `compose.live.yml` on `compose.yml` with four services: Caddy as
the proxy on `https://localhost:8443` with a certificate from Caddy's local authority, the
Prism365 app in production mode reachable only through the proxy, PostgreSQL 17 in the
`prism365_database` volume, and a one-time migration before the app starts.

Production mode needs the certificate, not a client secret. `npm run setup` prepares
everything `init` needs: the registration with the redirect URI
`https://localhost:8443/api/v1/auth/callback`, `secrets/entra-public.pem` uploaded to it,
`secrets/entra-private.pem`, and `.env`. Then:

```bash
scripts/docker-live.sh init
scripts/docker-live.sh up
```

`init` writes `.env.docker` (git-ignored, never copied into an image) with the tenant settings
from `.env` and newly generated database passwords and session encryption key. Keep that file:
the passwords and key belong to the data in the volume. `init` refuses to replace an existing
`.env.docker`; after running the setup again on an existing install, copy the changed
`ENTRA_CLIENT_ID`, `ENTRA_ADMIN_OBJECT_IDS` or `ENABLED_READ_FEATURES` into it yourself, then
run `down` and `up`. For a secrets folder other than `./secrets`, run
`PRISM_SECRETS_DIR=<dir> scripts/docker-live.sh init`.

The app container reads the certificate pair through a read-only mount as the image's `node`
user (user ID 1000). On Linux, if your account has another user ID, give that ID read access to
the key, for example with `setfacl -m u:1000:x secrets` and
`setfacl -m u:1000:r secrets/entra-private.pem`.

The browser warns about the local certificate until it is trusted;
`scripts/docker-live.sh export-ca` saves it and prints the macOS command to trust it.

Other commands: `down` stops everything and keeps the data, `logs` follows the app, migration
and proxy logs, `status` lists the containers, and `backup` writes a `pg_dump` to `backups/`
(git-ignored, readable only by you). A backup is written under a temporary name and fully read
back before it is kept, so a failed dump never looks like a good one. It contains tenant data in
plain form and retention does not reach it: keep it with `.env.docker`, on an encrypted disk,
out of cloud-synced folders, and delete old ones yourself. Never run `docker compose down -v`
against this project unless you intend to delete everything collected.

The app trusts one proxy hop (`TRUST_PROXY_HOPS=1`) so sign-in rate limits apply per client
address rather than to everything that comes through Caddy. On Docker Desktop every browser on
the Mac arrives from the same address, so they still share one limit there.

The image build refuses to run if the build context holds tenant data or keys (a PGlite
directory, a backup, a `.pem` or `.env` file) anywhere in the tree, not only at the top.

Collection is delegated: it runs on a signed-in administrator's session, which lasts eight
hours, and resumes on that administrator's next visit after the session ends or the server
restarts. Around-the-clock collection needs the separate application credential described in
ADR-0002.

## Self-host container

For the synthetic demo:

```bash
docker compose up --build -d app
```

Open `http://localhost:8080`. For live mode, configure HTTPS ingress and `PUBLIC_ORIGIN`, your
certificate files under the ignored `secrets/`, database passwords, and the encryption key.
Start the optional database and run the one-time migration before restarting the application:

```bash
docker compose --profile live up -d --wait db
docker compose --profile live run --rm migrate
docker compose up --build -d app
```

The application port binds to loopback. Terminate HTTPS in a trusted reverse proxy on the same
host. The application runs as a non-root user with a read-only filesystem, removed
capabilities, and no Docker socket. Its server serves both the built UI and `/api/v1` routes.

The base container images follow their Node and PostgreSQL major tags. Before a supported
release, pin reviewed image digests and test upgrades. Container building is checked in CI; no
deployment is performed by CI.

## Optional local AI

Configure `AI_LOCAL_MODEL` to an already installed Ollama model. The API calls only
`http://127.0.0.1:11434/api/chat`; redirects are rejected. When containerized, the optional
`ai` Compose profile shares the application network namespace so this endpoint remains
loopback:

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
Only PowerShell 7 packages with a fixed `.ps1` entrypoint, safe unique paths, an empty
network-host list, and a total canonical payload below 500000 bytes are accepted. See
[Lifecycle automation](AUTOMATION.md) for the full schema and a valid example.

Each upload receives a digest and unreviewed status. Neither import nor viewing the library
executes code. Package source is excluded from list responses, AI evidence, and public
repository content. Import only code you have permission to use and remove embedded secrets
first.

## Required pilot acceptance

Before enabling real organizational use, verify in a non-production tenant:

1. Allowed operator signs in; wrong tenant, operator, state or nonce is rejected.
2. User lookup succeeds; licenses and sign-ins independently explain missing access.
3. Pagination caps, empty results, throttling, and provider failures show honest coverage.
4. Removing an operator, signing out, or revoking consent blocks further access.
5. Sessions survive a server restart only when durable state is configured; no token is logged.
6. AI receives only selected minimized evidence and records failed attempts as well as success.
7. Package source remains private and cannot execute.

## Primary references

- [Add a certificate to an app using Microsoft Graph](https://learn.microsoft.com/en-us/graph/applications-how-to-add-certificate)
- [keyCredential resource type](https://learn.microsoft.com/en-us/graph/api/resources/keycredential?view=graph-rest-1.0)
- [Advanced query capabilities on Entra ID objects](https://learn.microsoft.com/en-us/graph/aad-advanced-queries)
- [MSAL Node token acquisition](https://learn.microsoft.com/en-us/entra/msal/javascript/node/acquire-token-requests)
- [Microsoft authorization code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Get user](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0)
- [List user license details](https://learn.microsoft.com/en-us/graph/api/user-list-licensedetails?view=graph-rest-1.0)
- [List sign-ins](https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0)
