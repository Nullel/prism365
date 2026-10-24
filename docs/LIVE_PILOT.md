# Live pilot: taking Prism365 out of demo

**Written:** 2026-10-22 against `main` at 89552ae. Baseline on this machine: `npm ci` done,
`npm run check` green (format, lint, typecheck, 121 tests; 6 PostgreSQL cases skipped because
no `DATABASE_ADMIN_URL`). PR #9 (`chore/publication-handoff`) is open and unmerged.

## 0. Status

**2026-10-22, evening:** Phase 1 reached. Live mode signed in against the pilot tenant
(delegated, single-tenant registration created in the portal by Noel, client secret for dev,
in-memory sessions). Remaining for Phase 1 acceptance: items 2 to 5 in §2.7, the Reports Reader
role check, and rotating the dev secret (it was pasted into a chat transcript during setup).

## 1. What "demo" is, in the code

- `DATA_MODE` defaults to `demo` (`apps/api/src/config.ts`). In demo the API has no identity
  service, `/v1/dashboard/summary` returns the synthetic `demoDashboard`, investigations return
  `createDemoUserInvestigation`, lifecycle previews are enabled, and the UI shows the
  "Exploring with synthetic data" banner, "Northstar Labs / Demo workspace" and "Demo" chips.
- `DATA_MODE=live` flips all of that: the identity service must be configured or the UI shows
  "Complete installation setup" with the missing keys; every `/v1/*` route except session,
  login, callback and health requires a session; `/v1/dashboard/summary` returns 409 and the
  Overview renders the "Tenant-wide reporting is next" placeholder; lifecycle previews return
  409; User Intelligence runs against Microsoft Graph with the operator's delegated token.

So "out of demo" is two different jobs:

1. **Live mode on a real tenant.** Removes every demo label, makes User Intelligence real.
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
gone. Open User Intelligence and look up a test user; the licence and sign-in sections each
report their own state.

### 2.5 Two code fixes before the pilot (GREEN, mech)

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
4. After sign-out, `/api/v1/investigations/user` answers 401.
5. No token or secret appears in the API console (Fastify logging is off; check anyway).

Items 6 and 7 of the original list (AI, packages) need PostgreSQL and belong to Phase 1b.

### 2.8 Phase 1b, optional: durable state

PostgreSQL 17 locally. Note that `compose.yml` publishes no host port for `db`, so to run
`npm run db:migrate` from the host either add `ports: ["127.0.0.1:5432:5432"]` to the `db`
service or run the `migrate` service inside Compose. Set `DATABASE_ADMIN_URL` and a 24+
character URL-safe `DATABASE_APP_PASSWORD`, migrate, then set `DATABASE_URL` for the
`prism365_app` login and a 64-hex `STATE_ENCRYPTION_KEY`. This unlocks AI grants (with
`AI_LOCAL_MODEL` and Ollama) and script import, and runs the six skipped PostgreSQL tests.

### 2.9 Phase 1c: the container (done 2026-10-22, see SETUP.md "Live install in Docker")

Production constraints apply: an HTTPS reverse proxy with a real origin, a certificate uploaded
to the registration and referenced by `ENTRA_CERTIFICATE_PATH` and `ENTRA_PRIVATE_KEY_PATH`,
PostgreSQL and the encryption key. Not needed to get out of demo on this Mac.

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
is still `reports.office.com` (a refusal names the host it saw) and how the file is encoded.

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
