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
   Microsoft asks you to consent to that tool. Consent for yourself only: leave Consent on
   behalf of your organization unticked, or every account in the tenant could use the tool with
   that permission. Microsoft keeps the consent after the run; see
   [Revoke the setup's access](#revoke-the-setups-access). The token itself stays in memory and
   is never written to disk.

2. Read the plan it prints, and type `yes` to apply it. It then:

   - registers a single-tenant web app (labelled `prism365-setup`, implicit grant off, with you as
     its owner) with the redirect URI `https://localhost:8443/api/v1/auth/callback` and the
     delegated Microsoft Graph permissions the chosen features need, and creates its service
     principal;
   - makes an RSA 2048 key and a self-signed certificate valid for one year, writes
     `secrets/entra-private.pem` (readable only by you, in a folder only you can open) and
     `secrets/entra-public.pem`, and adds the certificate to the registration;
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

6. Revoke the access you gave the Microsoft Graph command line tools (see below).

The setup never grants consent to Prism365's registration, never creates a client secret, never
asks for an application permission, and never removes a permission or a certificate unless you
ask (`--prune-permissions` for permissions). It is the only part of the project that changes
anything in Microsoft; Prism365 itself only reads.

### Which permissions it asks for

By default the registration asks for the core packs only: `directory.users.read`,
`user.licenses.read`, `tenant.licenses.read`, `audit.signins.read`, `audit.directory.read`,
`reports.authentication-methods.read`, `service.health.read` and `service.messages.read`. They
are what the People page, the audits, the sign-in audit, the change feed, service health and
Message center read, and each is a read permission a reporting tool is expected to hold.

The other packs (Conditional Access, risky users, Intune devices, enterprise applications, usage
reports and the usage-report privacy setting) reach further into the tenant, and the admin
consent link grants whatever the registration asks for. So they are asked for only when you
name them, as in `--features directory.users.read,identity.risk.read`, or with `--features all`.
Even `all` leaves out `directory.permission-grants.read`, because its permission,
`Directory.Read.All`, reads the whole directory; name it as well if you want delegated consents
shown (`--features all,directory.permission-grants.read`).

### Setup options

Pass them after `--`, as in `npm run setup -- --tenant contoso.onmicrosoft.com --dry-run`.

| Option | What it does |
|---|---|
| `--tenant <id or domain>` | The tenant to set up. Required. |
| `--name <display name>` | The registration's name. Default `Prism365`. |
| `--origin <origin>` | Where Prism365 runs. Default `https://localhost:8443`, the Docker install. The redirect URI is this origin plus `/api/v1/auth/callback`, and it becomes `PUBLIC_ORIGIN`. |
| `--dev` | Also registers the development server's redirect URI, `http://localhost:5173/api/v1/auth/callback`. |
| `--features <list>` | Comma-separated read packs to enable, or `all`. Default: the core packs above. |
| `--app-id <application ID>` | Use this existing registration. Without it, the setup uses the one `ENTRA_CLIENT_ID` in the settings file names, or creates a new one. |
| `--secrets-dir <dir>` | Where the certificate goes. Default `./secrets`. |
| `--env-file <file>` | The settings file to write. Default `.env`. |
| `--rotate-certificate` | Makes a new key and certificate even when a pair exists. |
| `--prune-permissions` | Removes requests for delegated permissions on an existing registration that the chosen features do not use. Consent already granted for them stays until you revoke it. |
| `--keep-other-permissions` | Keeps those delegated permissions instead. The admin consent link then grants them too. |
| `--keep-other-redirect-uris` | Keeps redirect URIs outside the chosen origins on an existing registration. |
| `--keep-other-certificates` | Keeps certificates this setup did not make on an existing registration or its service principal. |
| `--keep-other-owners` | Uses an existing registration that has owners besides you. They can change it at any time. |
| `--dry-run` | Signs in, reads, prints the plan and stops. Nothing is written anywhere. |
| `--yes` | Applies the plan without asking. |

### Running it again, and existing registrations

Every step checks what is already there first, so running the setup again is safe and finishes
a run that stopped halfway. A second run on a finished install changes nothing.

The setup never looks a registration up by its label or its name: anyone allowed to register
apps can give theirs the same ones. It works on the registration you name with `--app-id`, or
else on the one `ENTRA_CLIENT_ID` in the settings file names (the one this install already
uses), or else it creates a new one. It writes `ENTRA_CLIENT_ID` as soon as the registration
exists, so a run that stops after that is picked up again by the next one. If the settings file
names a registration the tenant does not have, it stops rather than make another.

Moving an install that already uses one registration to another with `--app-id` works the other
way round: the settings file keeps the old application ID until every step on the new one has
worked, so a run that fails leaves the install as it was. Once the move is done,
`secrets/setup-state.json` keeps the old application ID and the run prints the command that
switches back.

The `prism365-setup` label on a registration it creates is only a label. A create also carries a
second label unique to the run, so that when Microsoft answers a create with an error or not at
all, the setup can look for exactly that registration. It keeps looking for about 45 seconds,
the time a new registration can take to show, and creates again only if none has appeared by
then. A registration Microsoft takes even longer to show could still end up twice; the setup
then names the extra one at the end of the run, or names the label to look for if it stops. The
run label comes off when the run finishes.

Anyone who can read registrations sees that label once the registration exists and could copy it
onto one of theirs, so a registration that carries it is used only when it passes every check
below, with nothing kept. One that fails is not the one the run created: the setup stops, names
it, and neither uses it nor prints a consent link for it.

Before it changes an existing registration, it checks it, and stops without changing anything
if:

- you are not one of its owners, or it or its service principal has other owners, unless you
  pass `--keep-other-owners` (any owner can change it at any time);
- it or its service principal has a client secret, or it has a federated credential (anyone
  holding them could sign in as Prism365);
- it allows sign-in from mobile, desktop or single-page applications, or public client flows;
- it asks for, or has been granted, any application permission;
- consent has been given for delegated permissions Prism365 never asks for, or on another API;
- it asks for delegated permissions the chosen features do not use, unless you pass
  `--prune-permissions` or `--keep-other-permissions`;
- it sends sign-ins to an address outside the chosen origins, unless you pass
  `--keep-other-redirect-uris`;
- it or its service principal has a certificate this setup did not make, unless you pass
  `--keep-other-certificates`. When that certificate is the one its last rotation replaced, it
  says so.

Secrets, federated credentials, other sign-in platforms, application permissions and extra
consent have no option on purpose: remove them in the Entra admin center first. The setup never
removes consent itself, and `--prune-permissions` only removes requests: consent already given
stays until you revoke it there. When it updates a registration, the plan names each property it
sends; nothing else on the registration changes.

These checks describe the registration as it is when the setup runs. Its owners can change it at
any time afterwards, which is why the setup wants you to be its only owner.

The certificate's private key is never overwritten. If the key exists without its certificate,
the setup makes a certificate for that same key. If the two files do not belong together, the
key is not RSA 2048 or larger, or the certificate has expired, it stops before changing anything
and says so.

Adding a certificate to a registration replaces the registration's whole certificate list.
The setup reads the list again right before the update and stops if it differs from the list
the plan was made from, sends every existing certificate back unchanged with the new one, and
reads the list afterwards to check that each certificate it had just before is still there. If
Microsoft does not return a key in full, it stops without changing anything. Nobody should add
or remove the registration's certificates in the admin center while the setup runs: a change in
the moment between its last read and its update cannot be seen. (Microsoft's `addKey` would add
a certificate without replacing the list, but it needs proof signed by a certificate already on
the registration, which a first certificate cannot have, so the setup uses the update and these
checks for every case.)

### Rotating the certificate

The certificate lasts one year: Microsoft accepts a certificate credential for at most a year
from its start date when it is added through Microsoft Graph
([Add a certificate to an app](https://learn.microsoft.com/en-us/graph/applications-how-to-add-certificate)).
The setup checklist warns 30 days before it expires. To replace it:

```bash
npm run setup -- --tenant contoso.onmicrosoft.com --rotate-certificate
scripts/docker-live.sh down
scripts/docker-live.sh up
```

The new pair is written as `secrets/entra-private.new.pem` and `secrets/entra-public.new.pem`
and added to the registration first. Only once Microsoft has it does the setup move it into
place, keeping the old pair as `secrets/entra-private.previous.pem` and
`secrets/entra-public.previous.pem`. If anything fails before then, the working pair is
untouched and Prism365 keeps signing in with it. If the move itself is interrupted, the next run
finishes it.

The old certificate stays on the registration until you remove it under Certificates and
secrets in the Entra admin center. Once Prism365 signs in with the new one, remove the old
certificate there and delete the two `.previous` files; the setup will not rotate again while
they are there.

### File permissions on Windows

Windows has no file modes: the folder's access control list decides who can read the key, and
the setup leaves it as it is. Limit the secrets folder, and the files in it, to your account:

```powershell
icacls secrets /inheritance:r /grant:r "$($env:USERNAME):(OI)(CI)F"
```

### Revoke the setup's access

The setup signs in through the Microsoft Graph command line tools, and Microsoft keeps the
`Application.ReadWrite.All` consent you gave that tool after the run ends. Once setup is done,
remove it:

1. In the [Microsoft Entra admin center](https://entra.microsoft.com), open Entra ID, then
   Enterprise apps, then All applications, and select Microsoft Graph Command Line Tools.
2. Open Permissions. If you consented on behalf of your organization, the Admin consent tab
   lists `Application.ReadWrite.All`: choose it, select the menu next to it, and choose Revoke
   permission.
3. A consent for yourself only is on the User consent tab, which the admin center cannot
   revoke. Remove it with Microsoft Graph, for example in Graph Explorer signed in as a Cloud
   Application Administrator:
   `GET https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'&$select=id`
   gives the tool's ID; then
   `GET https://graph.microsoft.com/v1.0/oauth2PermissionGrants?$filter=clientId eq '<that ID>' and principalId eq '<your object ID>'`
   lists your grant. Its `scope` holds every permission you consented to for the tool, so take
   out only this one: send
   `PATCH https://graph.microsoft.com/v1.0/oauth2PermissionGrants/<grant ID>` with the body
   `{ "scope": "<the same list without Application.ReadWrite.All>" }`. If nothing else is left
   in it, `DELETE https://graph.microsoft.com/v1.0/oauth2PermissionGrants/<grant ID>` removes the
   grant instead.
4. Open Users and groups and remove your own assignment if you are listed. That alone does not
   remove the consent, which is why step 3 matters.

Microsoft documents these steps in
[Review permissions granted to enterprise applications](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/manage-application-permissions).
The setup prints a reminder at the end of every run.

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
- sessions and the access ledger survive restarts (only with PostgreSQL: `DATABASE_URL` and
  `STATE_ENCRYPTION_KEY`), and so does collected data (PostgreSQL, or `PRISM_DATA_DIR`);
- collection is running, or will resume;
- each audit has run at least once.

Loading the checklist only reads what Prism365 has stored. Check with Microsoft (a separate
request that carries the session's CSRF value, so no link or other site can trigger it) makes
one small read per granted pack: one item and only its ID where the endpoint allows it, a
filter that matches nothing where it allows only filters. It writes a `setup.probed` entry naming
the packs to the access ledger before it reads. The usage report lists every person, so it is
not read here; it is checked when its audit runs. When Microsoft refuses a read even though the
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
one year (365 days). The key must be a PEM file that starts with `-----BEGIN PRIVATE KEY-----`.

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
[Permissions](PERMISSIONS.md)). The source of truth is `packages/graph/src/permissions.ts`: a test checks the permission column
against it, and the role and licence columns summarise the notes kept there.

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
| `directory.roles.read` | `RoleManagement.Read.Directory` | Global Reader or Privileged Role Administrator | Eligible and time-bound roles need Entra ID P2 or ID Governance |
| `directory.groups.read` | `Group.Read.All` | Global Reader or Directory Readers | No extra licence |
| `directory.domains.read` | `Domain.Read.All` | Global Reader or Directory Readers | No extra licence |
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

Local development may use memory sessions, which reset on restart. AI requires PostgreSQL, and production live mode additionally requires encrypted durable state and
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
minutes and AI grants one hour. Expired state is
physically pruned on startup and every minute while the service runs. The runtime can insert and
read audit records but cannot update or delete them.

Script package import was removed on 2026-10-24. An install that imported packages before then
still holds them, encrypted, until they expire 30 days after import. The purge deletes them only
while the service runs in live mode with its identity configured, and only for that tenant, so
on any other install they stay. To remove them at once, run this as a PostgreSQL superuser
(row-level security hides the rows from other roles):

```sql
DELETE FROM prism365_state WHERE kind = 'automationPackage';
```

The access ledger keeps its import and removal entries either way.

Keep the encryption key backed up securely with the encrypted database backup. Losing the key
makes stored sessions unreadable. Key rotation, backup and restore, and
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
accounts by days since last sign-in, licenses to reclaim, MFA registration, guest accounts,
license capacity, Conditional Access policies and risky users. They use `directory.users.read`,
`audit.signins.read`, `tenant.licenses.read`, `reports.authentication-methods.read`,
`identity.conditional-access.read` and `identity.risk.read`. The first four are in the setup's
default list; the last two are asked for only when you name them or pass `--features all`.
Inactivity needs Microsoft Entra ID P1 or P2 and a role such as Reports Reader, Security Reader
or Global Reader; risky users need Entra ID P2. Results are kept (the last 10 successful runs
per audit) and can be exported as CSV. See `docs/audits.md`.

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
user (user ID 1000). On Linux, if your account has another user ID, give that ID access with an
access control list, including a default entry so files written later (a rotated pair) get it
too:

```bash
chmod 700 secrets
setfacl -m u:1000:x secrets
setfacl -d -m u:1000:r secrets
setfacl -m u:1000:r secrets/entra-private.pem secrets/entra-public.pem
```

The setup keeps that working: it leaves the group bits of a file or folder that has an access
list (or, where it cannot read the list, one whose group can read it), though it always takes
write access away from the group and from others. It writes new keys with mode 640 in a folder
with a default list so its entry applies, and after a rotation gives the new files the entries
the old ones had, listing each one in the plan first, or prints the `setfacl` command to run when
it cannot. A later run tries again only with the entries the `.previous` files carry;
`setup-state.json` notes only that something is left to do, never which entries.

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
AI can explain evidence and cannot change the tenant.

## Required pilot acceptance

Before enabling real organizational use, verify in a non-production tenant:

1. Allowed operator signs in; wrong tenant, operator, state or nonce is rejected.
2. User lookup succeeds; licenses and sign-ins independently explain missing access.
3. Pagination caps, empty results, throttling, and provider failures show honest coverage.
4. Removing an operator, signing out, or revoking consent blocks further access.
5. Sessions survive a server restart only when durable state is configured; no token is logged.
6. AI receives only selected minimized evidence and records failed attempts as well as success.

## Primary references

- [Add a certificate to an app using Microsoft Graph](https://learn.microsoft.com/en-us/graph/applications-how-to-add-certificate)
- [keyCredential resource type](https://learn.microsoft.com/en-us/graph/api/resources/keycredential?view=graph-rest-1.0)
- [Advanced query capabilities on Entra ID objects](https://learn.microsoft.com/en-us/graph/aad-advanced-queries)
- [MSAL Node token acquisition](https://learn.microsoft.com/en-us/entra/msal/javascript/node/acquire-token-requests)
- [Microsoft authorization code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Get user](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0)
- [List user license details](https://learn.microsoft.com/en-us/graph/api/user-list-licensedetails?view=graph-rest-1.0)
- [List sign-ins](https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0)
