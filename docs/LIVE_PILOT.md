# Live pilot: taking Prism365 out of demo

**Written:** 2026-10-22 against `main` at 89552ae. Baseline on this machine: `npm ci` done,
`npm run check` green (format, lint, typecheck, 121 tests; 6 PostgreSQL cases skipped because
no `DATABASE_ADMIN_URL`). PR #9 (`chore/publication-handoff`) is open and unmerged.

**Refreshed:** 2026-11-17 against `main` at b9f7fd3: §2.10 (the full audit suite) added, with
its tables generated from the code by `npm run pilot:matrix`. Sections 3 to 5 were not
re-checked in this refresh and may be out of date.

## 0. Status

**2026-10-22, evening:** Phase 1 reached. Live mode signed in against the pilot tenant
(delegated, single-tenant registration created in the portal by Noel, client secret for dev,
in-memory sessions). Remaining for Phase 1 acceptance: items 2 to 5 in §2.7, the Reports Reader
role check, and rotating the dev secret (it was pasted into a chat transcript during setup).

**2026-11-17:** the Phase 1 items above are still open. The next step, running every audit and
collector on live data, is Phase 1d (§2.10). Rotate the dev secret before starting it.

## 1. What "demo" is, in the code

- `DATA_MODE` defaults to `demo` (`apps/api/src/config.ts`). In demo the API has no identity
  service, `/v1/dashboard/summary` returns the synthetic `demoDashboard`, and the UI shows the
  "Exploring with synthetic data" banner, "Northstar Labs / Demo workspace" and "Demo" chips.
- `DATA_MODE=live` flips all of that: the identity service must be configured or the UI shows
  "Complete installation setup" with the missing keys; every `/v1/*` route except session,
  login, callback and health requires a session; `/v1/dashboard/summary` returns 409 and the
  Overview renders the "Tenant-wide reporting is next" placeholder; People looks accounts up in
  Microsoft Graph with the operator's delegated token.

So "out of demo" is two different jobs:

1. **Live mode on a real tenant.** Removes every demo label, makes People lookups real.
   This is Phase 1 and it is mostly operator work plus two small code fixes.
2. **A real Overview.** Needs tenant-wide collection, which does not exist yet (roadmap
   Milestones 2 and 3). Phase 2 scopes the first honest version.

## 2. Phase 1: live mode on a tenant you control

Use the dev path, not the container. The Dockerfile sets `NODE_ENV=production`, and in
production `loadConfig` demands an HTTPS `PUBLIC_ORIGIN`, a certificate credential and
`DATABASE_URL` plus `STATE_ENCRYPTION_KEY`. The dev path (`npm run dev:api` and
`npm run dev:web`) accepts `http://localhost:5173`, a client secret and the in-memory store.
That is the shortest route to "Microsoft connected".

### 2.1 Pick the tenant

| Option | Cost | What works | Notes |
|---|---|---|---|
| **A. Your own Entra tenant** (recommended) | free | directory, licence details, sign-ins once the Entra ID P2 trial is activated | You are Global Admin, so registration and consent take minutes. No Microsoft 365 workloads, so usage reports would be empty. Verify the P2 trial is still offered without a card at activation time. |
| B. Microsoft 365 E5 or Business Premium trial tenant | card required, one month, 25 seats (Microsoft's page) | everything, including P1/P2 and real mail, Teams, SharePoint | Only worth it when Phase 2 needs workload data. |
| C. A tenant you do not own | free | same as A | Production tenant. The Global Admin must create the registration and consent; your account needs Reports Reader, Security Reader or Global Reader for sign-ins and licence details. `SETUP.md` says start non-production. Keep for later. |

Option A is enough for Phase 1. The full audit suite (§2.10) needs Option B: several audits read
Intune, Entra ID P2 or Microsoft 365 workload data that an Entra-only tenant does not have.

The Microsoft 365 Developer Program sandbox is no longer open sign-up; community reports in
2026 say it is gated to Visual Studio Pro/Enterprise subscribers and partner programs.

### 2.2 App registration (values the code requires)

- Supported account types: **this organizational directory only**. `config.ts` validates the
  tenant ID and `identity.ts` checks the token's `tid` against it.
- Platform: **Web**, not SPA. Redirect URI exactly `http://localhost:5173/api/v1/auth/callback`.
  Leave implicit grant off; the server redeems the code with PKCE.
- API permissions, delegated: `User.Read`, `User.Read.All`, `LicenseAssignment.Read.All`,
  `AuditLog.Read.All`. Grant admin consent so the per-feature consent links do not hit the
  "needs admin approval" wall.
- Credential: a client secret for the dev path (short expiry). A certificate is only needed
  for the container path.
- Record three values: Directory (tenant) ID, Application (client) ID, and the Object ID of
  the account that will operate Prism365.

**Do it from the terminal instead of the portal.** The CLI for Microsoft 365 (installed on this
Mac as `m365`, v11.11.0) creates the registration with exactly these settings in one command.
It needs its own app registration in the tenant first; `m365 setup` creates that interactively
(broad permissions, acceptable in a tenant you own for testing). Then:

```bash
m365 login
```

```bash
m365 entra app add --name "Prism365 (dev)" --platform web --redirectUris "http://localhost:5173/api/v1/auth/callback" --apisDelegated "https://graph.microsoft.com/User.Read,https://graph.microsoft.com/User.Read.All,https://graph.microsoft.com/LicenseAssignment.Read.All,https://graph.microsoft.com/AuditLog.Read.All" --grantAdminConsent --withSecret --output json
```

The JSON output carries `appId`, `tenantId` and the generated secret (named `Default`, one year).
Put the secret straight into `.env` and nowhere else. Then the operator's object ID:

```bash
m365 entra user get --id "@meId" --properties "id,userPrincipalName"
```

### 2.3 `.env`

Copy `.env.example` to `.env` (git-ignored) and set:

```
HOST=127.0.0.1
PORT=8787
DATA_MODE=live
PUBLIC_ORIGIN=http://localhost:5173
MICROSOFT_CLOUD=worldwide
ENTRA_TENANT_ID=<tenant uuid>
ENTRA_CLIENT_ID=<client uuid>
ENTRA_ADMIN_OBJECT_IDS=<operator object id>
ENABLED_READ_FEATURES=directory.users.read,user.licenses.read,audit.signins.read
ENTRA_CLIENT_SECRET=<dev secret>
```

All three IDs must be UUIDs or `loadConfig` throws. Leave `DATABASE_URL` unset for now: the
memory store works, sessions just die on every restart (including `tsx watch` restarts).
Leave `AI_LOCAL_MODEL` unset; AI needs PostgreSQL anyway.

### 2.4 Run and sign in

```bash
npm run dev:api
```

```bash
npm run dev:web
```

Open `http://localhost:5173`, not `127.0.0.1`: the cookie and the origin check are against
`PUBLIC_ORIGIN`. Expected: topbar says "Live setup", the Tenant connection card offers
"Sign in with Microsoft". After the round trip: "Microsoft connected", "Signed in as …", the
workspace switcher reads "Your organization / Microsoft 365 workspace", the demo banner is
gone. Open People and look up a test user by their full work email; the licence and sign-in
sections each report their own state.

### 2.5 Two code fixes before the pilot (GREEN)

1. **Surface a failed sign-in.** Done 2026-10-22 (`readAuthNotice` in `App.tsx`, category log in `app.ts`). Both failure paths in `/v1/auth/callback` redirect to
   `/?auth=failed`, and nothing in `apps/web` reads that parameter, so a rejected operator
   (wrong tenant, not on the allowlist, expired state) lands back on the connection card with
   no message. Read `auth=failed` in `App.tsx` and show one sentence on the card. Log a
   redacted reason server-side (category only, no token, no object ID) so the pilot can tell
   allowlist failures from token failures.
2. **Fix the repo links.** Done 2026-10-22. The GitHub account itself was renamed from `nlj3`
   to `Nullel` (the old login returns 404), so CODEOWNERS, GOVERNANCE.md, PUBLISHING.md, the
   issue-template contact link and the sidebar link now point at `Nullel`. PR #9 still carries
   the old name in README, SECURITY.md and TASKS.md; fix those when it merges.

### 2.6 Gotchas found in the code

- If `localhost` resolves to `::1` on this Mac the Vite server (bound to 127.0.0.1) will not
  answer. Then set `PUBLIC_ORIGIN=http://127.0.0.1:5173` and change the redirect URI to match;
  Entra accepts loopback HTTP.
- Login starts are rate-limited to 10 per minute per IP. OAuth state lives 5 minutes, a
  session 8 hours.
- `ENTRA_ADMIN_OBJECT_IDS` is a hard allowlist checked at callback and on every session read.
- Sign-ins need Entra ID P1 or P2 on the tenant and a reader role on the operator; a 403 shows
  as the "unknown" state with the honest message, by design.
- `ENTRA_CLIENT_SECRET` is ignored when `NODE_ENV=production`; the key silently joins
  `missingConfiguration` and the UI shows "Complete installation setup".
- The in-memory store caps at 1,000 entries and 10,000 audit rows. Fine for a pilot.

### 2.7 Acceptance for Phase 1

From the "Required pilot acceptance" list in `docs/SETUP.md`, trimmed to what the dev path can
prove:

1. The allow-listed operator signs in; an account not on the list is rejected and the UI says so.
2. User lookup succeeds; with a feature removed from `ENABLED_READ_FEATURES` and the API
   restarted, that section explains missing access on its own.
3. A missing user returns the 404 state; an empty sign-in window says so without implying
   inactivity.
4. After sign-out, `/api/v1/people/lookup?email=...` answers 401.
5. No token or secret appears in the API console (Fastify logging is off; check anyway).

Item 6 of the original list (AI) needs PostgreSQL and belongs to Phase 1b.

### 2.8 Phase 1b, optional: durable state

PostgreSQL 17 locally. Note that `compose.yml` publishes no host port for `db`, so to run
`npm run db:migrate` from the host either add `ports: ["127.0.0.1:5432:5432"]` to the `db`
service or run the `migrate` service inside Compose. Set `DATABASE_ADMIN_URL` and a 24+
character URL-safe `DATABASE_APP_PASSWORD`, migrate, then set `DATABASE_URL` for the
`prism365_app` login and a 64-hex `STATE_ENCRYPTION_KEY`. This unlocks AI grants (with
`AI_LOCAL_MODEL` and Ollama) and runs the six skipped PostgreSQL tests.

### 2.9 Phase 1c: the container (done 2026-10-22, see SETUP.md "Live install in Docker")

Production constraints apply: an HTTPS reverse proxy with a real origin, a certificate uploaded
to the registration and referenced by `ENTRA_CERTIFICATE_PATH` and `ENTRA_PRIVATE_KEY_PATH`,
PostgreSQL and the encryption key. Not needed to get out of demo on this Mac.

### 2.10 Phase 1d: the full audit suite

Phase 1 proves sign-in and People lookups. Since this document was written, Prism365 has gained
on-demand audits, four collectors, alerts and handoffs. Running all of them on live data is the
step out of pre-alpha. The tables at the end of this section are generated from the code: run
`npm run pilot:matrix -- --write` after adding an audit, a collector or a read pack, and
`npm run pilot:matrix -- --check` to confirm they are current.

**First, rotate the dev client secret** from §0. It was pasted into a chat transcript, so treat
it as exposed: add a new secret to the registration, put it in `.env`, then delete the old one.

**Tenant.** Use Option B from §2.1. Option A leaves several audits empty: Managed devices needs
an Intune license; Risky users, and the eligible assignments in Privileged roles, need Entra ID
P2; Microsoft 365 app usage, Mailbox storage, the usage columns of Groups and Teams and the
workload columns of Account activity need real Exchange, SharePoint and Teams activity. At
sign-up, check on Microsoft's trial page that the trial still includes Entra ID P2 and Intune.

**Operator roles**, all read-only, taken from the requirements in the tables below:

| Role | Why |
|---|---|
| Global Reader | Most audits name it, including Conditional Access, Risky users, Managed devices and Privileged roles. |
| Reports Reader | The sign-in and directory-change collectors require it (`role.reportsReader`), and the per-user usage reports need it: Global Reader sees tenant totals only. |
| Directory Readers | App consents reads app role assignments, and Microsoft does not list Global Reader for that. |

**Report names.** The usage reports conceal names while the tenant's report privacy setting says
so; a Global Administrator controls that setting. The usage audits read it with
`ReportSettings.Read.All` and say when names are concealed (`docs/audits.md`). That pack,
`reports.settings.read`, is optional in the tables below; enable it for the pilot so acceptance
item 2 can be checked. Then show names, or expect the usage audits to list concealed identifiers
instead of people.

**Acceptance for Phase 1d.** With PostgreSQL configured (§2.8), `ENABLED_READ_FEATURES` set to
the full-suite line below plus `reports.settings.read`, and the operator holding the three roles:

1. No audit on the Audits page says "Needs …" or "Not enabled here".
2. Each audit runs to completion on live data. Where a result is partial (no P2, no workload
   activity, concealed names), the run says so in its own words; an empty table never reads as a
   clean result.
3. Combined risks runs after the audits it joins and uses no result older than seven days.
4. The sign-in and directory-change collectors complete a first pass, and a short API restart
   leaves no gap finding once collection resumes. Service health and Message center import the
   current items.
5. One alert rule delivers to a webhook or SMTP address you control, as `docs/alerts.md`
   describes.
6. Export for action on an audit that offers it produces a handoff. A receipt written by hand for
   that handoff, every item `skipped`, imports, and importing it again keeps it once
   (`docs/handoff.md` §5). This proves the loop before any action tool exists.
7. No token, secret or client credential appears in the API console or logs (§2.7, item 5).

#### Generated requirements

<!-- pilot-matrix:start -->
#### Audits (16)

| Audit | Read packs | Permissions | Optional permissions | Roles and licences |
|---|---|---|---|---|
| Inactive accounts | `directory.users.read`, `audit.signins.read` | `User.Read.All`, `AuditLog.Read.All` | `LicenseAssignment.Read.All` | Microsoft Entra ID P1 or P2 in the tenant, which Microsoft requires for sign-in dates. A role that can read sign-in activity, such as Reports Reader, Security Reader or Global Reader. |
| Licenses to reclaim | `directory.users.read` | `User.Read.All` | `AuditLog.Read.All`, `LicenseAssignment.Read.All` | Without sign-in dates (AuditLog.Read.All with Microsoft Entra ID P1 or P2), only disabled accounts are listed. |
| MFA registration | `reports.authentication-methods.read` | `AuditLog.Read.All` | none | A role such as Reports Reader, Security Reader, Security Administrator or Global Reader. |
| Guest accounts | `directory.users.read` | `User.Read.All` | `AuditLog.Read.All`, `LicenseAssignment.Read.All` | Sign-in dates need AuditLog.Read.All and Microsoft Entra ID P1 or P2. Without them, only invitation states are shown. |
| License capacity | `tenant.licenses.read` | `LicenseAssignment.Read.All` | none | A role that can read subscriptions, such as Directory Readers, Global Reader or License Administrator. |
| Conditional Access policies | `identity.conditional-access.read` | `Policy.Read.All` | `User.Read.All`, `LicenseAssignment.Read.All` | A role such as Global Reader, Security Reader, Security Administrator or Conditional Access Administrator. Microsoft Entra ID P1 or P2. |
| Risky users | `identity.risk.read` | `IdentityRiskyUser.Read.All` | none | Microsoft Entra ID P2. A role such as Global Reader, Security Reader, Security Operator or Security Administrator. |
| Managed devices | `devices.intune.read` | `DeviceManagementManagedDevices.Read.All` | none | An active Intune license in the tenant, which Microsoft requires for its device API. A role that can read Intune devices, such as Global Reader, Security Reader, Helpdesk Administrator or Intune Administrator, or the Read Only Operator role in Intune. |
| App consents | `directory.applications.read` | `Application.Read.All` | `Directory.Read.All` | A role that can read applications and their app role assignments, such as Directory Readers, Application Administrator or Cloud Application Administrator. Microsoft does not list Global Reader for reading app role assignments. Delegated permissions (optional) need Directory.Read.All and a role such as Global Reader or Directory Readers. |
| Microsoft 365 app usage | `reports.usage.read` | `Reports.Read.All` | `ReportSettings.Read.All` | Reports Reader or a service administrator role (Exchange, SharePoint or Teams). Global Reader and Usage Summary Reports Reader see tenant totals only, not this per-user report. Names show only when the tenant's report privacy setting shows them. A Global Administrator controls that setting. |
| Account activity | `directory.users.read`, `audit.signins.read` | `User.Read.All`, `AuditLog.Read.All` | `Reports.Read.All`, `ReportSettings.Read.All`, `LicenseAssignment.Read.All` | Microsoft Entra ID P1 or P2 in the tenant, which Microsoft requires for sign-in dates. A role that can read sign-in activity, such as Reports Reader, Security Reader or Global Reader. The workload columns need Reports Reader or a service administrator role (Exchange, SharePoint or Teams). Global Reader sees tenant totals only. Names show in the usage reports only when the tenant's report privacy setting shows them. A Global Administrator controls that setting. |
| Privileged roles | `directory.roles.read` | `RoleManagement.Read.Directory` | `User.Read.All`, `Application.Read.All`, `AuditLog.Read.All` | A role that can read directory roles, such as Global Reader, Directory Readers or Privileged Role Administrator. Directory Readers cannot read Privileged Identity Management schedules. Microsoft Entra ID P2 or ID Governance for eligible and time-bound assignments. Without it only active assignments are listed. |
| Passwords | `directory.users.read` | `User.Read.All` | `Domain.Read.All` | Expiry dates need the domains read pack and a role such as Global Reader or Directory Readers. Without it the audit shows password age and the account settings only. |
| Groups and Teams | `directory.groups.read` | `Group.Read.All` | `Reports.Read.All`, `ReportSettings.Read.All` | A role that can read groups and their owners, such as Global Reader or Directory Readers. The usage columns need the usage reports read pack and a role that can read them, such as Reports Reader or Teams Administrator. |
| Mailbox storage | `reports.usage.read` | `Reports.Read.All` | `ReportSettings.Read.All` | A role that can read the usage reports, such as Reports Reader or Exchange Administrator. Names show in the usage reports only when the tenant's report privacy setting shows them. A Global Administrator controls that setting. |
| Combined risks | none | none | `Policy.Read.All` | Run the audits it joins first: Account activity, MFA registration, Privileged roles, Risky users, Guest accounts, Passwords, Managed devices, Mailbox storage and Groups and Teams. A result older than seven days is not used. Conditional Access policies are read when the Conditional Access read pack is enabled and granted. Without them only MFA registration is checked. |

#### Collectors (4)

| Module | Read packs | Permissions | Requires |
|---|---|---|---|
| Sign-in audit | `audit.signins.read` | `AuditLog.Read.All` | `entra.p1`, `role.reportsReader` |
| Service health | `service.health.read` | `ServiceHealth.Read.All` | none |
| Message center | `service.messages.read` | `ServiceMessage.Read.All` | none |
| Directory changes | `audit.directory.read` | `AuditLog.Read.All` | `role.reportsReader` |

#### `ENABLED_READ_FEATURES` for the full suite

Every pack an audit or collector requires, plus the packs other pages read (`user.licenses.read`):

```
ENABLED_READ_FEATURES=audit.directory.read,audit.signins.read,devices.intune.read,directory.applications.read,directory.groups.read,directory.roles.read,directory.users.read,identity.conditional-access.read,identity.risk.read,reports.authentication-methods.read,reports.usage.read,service.health.read,service.messages.read,tenant.licenses.read,user.licenses.read
```

Optional only (an audit runs without them and shows fewer columns): `directory.domains.read`, `directory.permission-grants.read`, `reports.settings.read`. Add one only for what it adds.

#### Delegated permissions for the registration

`Application.Read.All`, `AuditLog.Read.All`, `DeviceManagementManagedDevices.Read.All`, `Group.Read.All`, `IdentityRiskyUser.Read.All`, `LicenseAssignment.Read.All`, `Policy.Read.All`, `Reports.Read.All`, `RoleManagement.Read.Directory`, `ServiceHealth.Read.All`, `ServiceMessage.Read.All`, `User.Read`, `User.Read.All`.

```bash
m365 entra app add --name "Prism365 (pilot)" --platform web --redirectUris "http://localhost:5173/api/v1/auth/callback" --apisDelegated "https://graph.microsoft.com/Application.Read.All,https://graph.microsoft.com/AuditLog.Read.All,https://graph.microsoft.com/DeviceManagementManagedDevices.Read.All,https://graph.microsoft.com/Group.Read.All,https://graph.microsoft.com/IdentityRiskyUser.Read.All,https://graph.microsoft.com/LicenseAssignment.Read.All,https://graph.microsoft.com/Policy.Read.All,https://graph.microsoft.com/Reports.Read.All,https://graph.microsoft.com/RoleManagement.Read.Directory,https://graph.microsoft.com/ServiceHealth.Read.All,https://graph.microsoft.com/ServiceMessage.Read.All,https://graph.microsoft.com/User.Read,https://graph.microsoft.com/User.Read.All" --grantAdminConsent --withSecret --output json
```

For each optional pack you enable, append its permission to `--apisDelegated`: `https://graph.microsoft.com/Directory.Read.All`, `https://graph.microsoft.com/Domain.Read.All`, `https://graph.microsoft.com/ReportSettings.Read.All`.
<!-- pilot-matrix:end -->

## 3. Phase 2: make the Overview live (YELLOW, about two days)

After Phase 1 the Overview is still a placeholder. The existing `dashboardSummarySchema`
(metrics, sign-in trend, recent activity, time window) can be filled on request from the
operator's delegated session, with no scheduler and no new storage:

| Metric | Source | Pack | Status |
|---|---|---|---|
| Guest accounts | `users?$filter=userType eq 'Guest'&$count=true` with `ConsistencyLevel: eventual` | `directory.users.read` | path already allowed by the client |
| Seats consumed vs enabled | `subscribedSkus` | `tenant.licenses.read` | exists in `permissions.ts`; must be added to `supportedReadFeatures`, the login enum in `app.ts`, and `TenantConnection.tsx` |
| Protected sign-ins % and 7-day trend | `auditLogs/signIns` for the window, aggregated per day by status, risk and Conditional Access outcome | `audit.signins.read` | bounded by the client's page and record caps; report partial coverage honestly |
| Open service incidents | `admin/serviceAnnouncement/healthOverviews` and `issues` | `service.health.read` | add the feature and consent `ServiceHealth.Read.All` |
| Recent activity | open service issues plus failed-sign-in spikes | the two above | |

Usage reports (`Reports.Read.All`): most `reports/get*` functions answer 302 with a
pre-authenticated CSV link on `reports.office.com`. `GraphHttpClient.request()` still refuses
every 3xx; `getReport()` is the one reader that follows it (ADR-0010), for eight per-user
reports by period, without sending the token, and parses the file with a bounded CSV reader. The
per-user Microsoft 365 apps report (`getM365AppUserDetail`) also has a documented JSON answer,
which the Microsoft 365 app usage audit reads. On the pilot tenant, check that the download host
is still `reports.office.com` and how the file is encoded. If Microsoft has moved the download, the
audit page says "Microsoft sent the report from a host Prism365 does not recognise" with the host,
and the server log line carries the same host with the `report_host` category; the link itself is
never shown or logged.

Deliverable: `/v1/dashboard/summary` returns live data in live mode, `MetricCard` reads
`metric.state` instead of the hard-coded "Demo" chip, and the placeholder section goes away.
Tests use the existing fake-fetch pattern in `apps/api/test`.

## 4. Phase 3: persisted collection

Roadmap Milestone 2: scheduled jobs on an application credential, snapshot tables, daily diffs,
trends. The signal table and derived layer are already designed in
`~/Documents/m365-skynet/docs/SPEC.md` and fold in here. Not part of "out of demo".

## 5. Housekeeping noticed on the way

- PR #9 adds the CI smoke test and the `pg` pool error handler; merge or close it before pilot
  changes land, since it also touches `server.ts` and `App.tsx`.
- Three Dependabot PRs (react, react-dom, vitest) and one for `actions/checkout`.
- `npm ci` warned that esbuild and fsevents install scripts are not yet in npm 11's
  `allowScripts`; the check suite still passed.
