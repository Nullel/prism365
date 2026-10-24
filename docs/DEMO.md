# The portfolio demo

Status: built. Files: `scripts/demo/tenant.ts` (the invented Microsoft tenant),
`scripts/demo/record.ts` (the recorder), `apps/web/src/demo/` (the replay shim and the banner),
`apps/web/demo/fixtures.json` (the recording), `apps/web/test/demo-replay.test.ts` (its tests).

The demo is a static copy of the web app that a visitor can click through with no server behind
it. Every screen shows data, and that data is not made up by the page. It is what the real
Prism365 API returned when it ran against an invented company.

## What is in it, and what is real

The company is Larkspur Analytical, an analytical laboratory with about fifty accounts: a
laboratory in Madison, an office in Minneapolis, a few remote staff, four guests, two
administrators and a shared mailbox. All of it is invented. Nothing in the demo comes from a
real tenant, a real person or a real organisation.

- Every domain ends in `.example`, which RFC 2606 reserves so it can never belong to anyone.
- Every IP address sits in a documentation range: 192.0.2.0/24, 198.51.100.0/24,
  203.0.113.0/24, or 2001:db8::/32.
- Phone numbers use the 555-0100 to 555-0199 block that North American numbering keeps for
  fiction.
- The Microsoft service-health incidents and Message center posts were written for this demo.
  They look like Microsoft's, but Microsoft never published them.

What is real is the Prism365 code. The recorder runs the actual API, the actual collectors, the
actual detection rules and the actual audits against a synthetic copy of the Microsoft Graph, so
the findings, counts, tables and notifications are what Prism365 really produces for that data.

Planted in the sign-in history, and found by the real rules, are four stories to look for:

- An impossible-travel sign-in from Romania, about an hour after the same account signed in at
  its Madison laboratory (a critical finding).
- A password spray: one address trying one password against seven accounts, eight minutes apart
  (acknowledged).
- Five failed MFA prompts against one account in about six minutes.
- A successful sign-in right after five failed attempts, from the account's own laptop (the
  failed attempts are resolved, the success is still open).

The directory has accounts that went quiet 8 to 400 days ago, never-signed-in accounts, two
disabled accounts that still hold licenses, a license assigned by a group, a guest invitation
that has waited 45 days, an administrator with no MFA, and subscriptions with no seats left, with
a grace period, with spare seats and with a free product that has 10,000.

## Record it

```
npm run demo:record
```

This takes a few seconds. It starts the API inside the process (nothing listens on a port), signs
in a fake administrator, starts collection so the real collectors read the invented Graph, waits
until the sign-in history is loaded and service health and Message center have been checked, runs
every audit, then asks the API for every response the web app uses and writes them to
`apps/web/demo/fixtures.json`. It never contacts Microsoft, never uses a database server and
never touches Docker.

The recording is repeatable: the same people, devices and stories come out every time. Only the
timestamps move, because they are measured back from the moment of recording.

Before it writes anything, the recorder checks the whole file. Every email address, every link
host and every IP address in it must belong to the invented tenant, to a `.example` domain, to a
documentation range, or to Microsoft's own admin center. If something else turns up, nothing is
written.

Run it again whenever the API, a module, an audit or the contracts change in a way a screen can
see, and commit the new file. It is about 1.6 MB. Biome skips files over 1 MiB and prints a
warning about this one; that is expected.

## Build it

```
npm run build:demo
```

This builds `apps/web/dist-demo/`: the normal app plus a small shim that answers `/api/v1/...`
from the recording. Serve that folder with any static server, for example:

```
python3 -m http.server 4173 --directory apps/web/dist-demo
```

The build uses relative paths, so the folder also works under a sub-path on a static host.
Nothing is hosted or deployed by this repository.

The normal `npm run build` is not affected. The shim, the recording and the banner are loaded
only when the build sets `VITE_DEMO=true` (see `apps/web/.env.demo`), so they are not in
`apps/web/dist`.

Biome has no ignore for `apps/web/dist-demo`, so `npm run check` fails while that folder exists.
Add `"!**/dist-demo"` to `files.includes` in `biome.json`, or delete the folder before checking.

## How the demo behaves

- Reads come from the recording. When a request has no recording, the page gets a 404 saying
  that this is not available in the demo.
- Dates are moved forward by a whole number of days, so the demo always looks current. The time
  of day is kept, and nothing is ever in the future.
- Writes are simulated in memory and reset when the page reloads. Marking a notification
  reviewed, done, muted or read, changing a finding's state, saving notification settings,
  starting and stopping collection, and running an audit again all work. A new audit run shows
  as running and finishes two seconds later with the recorded rows.
- The CSV export is built in the browser from the recorded rows, with the same filters and sort
  the page shows.
- The people search works on the recorded list. Looking someone up by email, and the Live lookup
  page, only know the accounts in the recording, for example `admin@larkspur.example`,
  `ingrid.solheim@larkspur.example`, `dagny.holloway@larkspur.example`,
  `kasimir.orne@larkspur.example` and `wendell.aris@larkspur.example`.
- Sign-out does nothing, and the links that start a Microsoft sign-in are left as they are.
- AI analysis and package import are shown as unavailable, because the recorded installation
  has neither.

## Keeping it honest

The banner at the top of every page says the company is invented and that the screens are
recorded output. That is the whole claim, and it should stay true. If a screen needs data the
real code does not produce, change the invented tenant in `scripts/demo/tenant.ts` so the real
code produces it. Do not edit the recording by hand. Prism365 does not invent security evidence,
and the demo does not either: it shows what the product does with a company that never existed.

## Publishing

The demo is published to GitHub Pages from `main` by `.github/workflows/demo.yml`, which builds it
with `npm run build:demo` whenever the web app or the shared contracts change, and can also be run
by hand. It is served at https://nullel.github.io/prism365/.
