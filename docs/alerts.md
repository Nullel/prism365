# Alerts

Alerts tell an administrator that something needs attention without them opening Prism365:
a new or escalated finding, a directory change that matters, collection that stopped, or an
audit number that crossed a line. They can also send an audit's scheduled results and a weekly
summary. The page is Settings, Alerts.

Prism365 still never writes to Microsoft, and it never sends mail through Microsoft either.
Email goes through the operator's own mail relay; everything else goes to a webhook the
administrator names. ADR-0011 explains why, and what leaves the server.

## Destinations

A destination is where alerts go: a list of email addresses, or one webhook. A tenant can have
20. Each has an "include details" switch, off by default (see What an alert holds).

A destination a rule, a scheduled export or the weekly summary still sends to cannot be deleted;
change those first. There is no way to edit a destination: delete it and add it again.

### Email

Email needs these settings on the server. None of them can be set or read from the page.

| Setting | What it holds |
|---|---|
| `SMTP_URL` | `smtps://relay.example:465` for TLS from the start, or `smtp://relay.example:587`, which then must upgrade with STARTTLS. A user may be named in it (`smtps://alerts@relay.example`); a password never. No path, query or other option is accepted. |
| `SMTP_PASSWORD_FILE` | A file with the relay password on one line, read once at start. Required when `SMTP_URL` names a user. `SMTP_PASSWORD` is refused, so the password never sits in the environment. |
| `SMTP_FROM` | The sender, one plain address. |
| `ALERT_EMAIL_DOMAINS` | The domains recipients must be in, comma-separated. A recipient must be in one exactly: `ops.example.invalid` is not `example.invalid`. |

In the Docker install, `scripts/docker-live.sh init` copies `SMTP_URL`, `SMTP_FROM`,
`ALERT_EMAIL_DOMAINS` and `WEBHOOK_ALLOWED_PORTS` from your `.env`; when `SMTP_URL` names a user, put
the password in `secrets/smtp-password` first, which is mounted read-only at
`/run/secrets/smtp-password`.

Plain SMTP is refused, and so is a relay certificate that does not verify (TLS 1.2 or later).
Recipients are checked when a destination is saved and again before every message, so taking a
domain off the list stops mail to it at once. One recipient outside the list stops the whole
message. Each address is one plain mailbox: ASCII only, no display name or comment, and a part
before the @ of letters, digits and `. _ + ' -` with no dot at either end or two together. That
leaves out `%`, `!` and quoted local parts, which a relay may read as a route to another host
(`ops%attacker.example@example.invalid`), so an address can only mean a mailbox in its own domain.
An email destination saved before a recipient became refused (a domain taken off the list) is
marked on the page and gets nothing.

Messages are plain text, from "Prism365", with `Auto-Submitted: auto-generated`. The subject and
any attachment's name never carry a line break, whatever the provider text in them.

### Webhooks

Three formats:

- Teams: the Workflows trigger "When a webhook request is received". The message is one
  Adaptive Card of text blocks with a single action, a link back to Prism365.
- Slack: an incoming webhook, `{"text": ...}` in plain text, with link previews off.
- Generic: the Prism365 JSON below, for anything else.

The address must be HTTPS with a host name: no address literal, no user name or password, and
port 443 unless the operator lists another in `WEBHOOK_ALLOWED_PORTS` (comma-separated). Local
names (`localhost`, `.local`, `.internal`, `.arpa` and the like) are refused, and so are
addresses with whitespace, control characters or backslashes, which URL parsers drop or read
differently.

The host is looked up when each message is sent, not when the destination is saved. If any
address it resolves to is in a refused range, nothing is sent:

- IPv4: `0.0.0.0/8`, `10.0.0.0/8`, `100.64.0.0/10` (CGNAT), `127.0.0.0/8`, `169.254.0.0/16`
  (link-local and the 169.254.169.254 metadata service), `172.16.0.0/12`, `192.0.0.0/24`,
  `192.0.2.0/24`, `192.88.99.0/24`, `192.168.0.0/16`, `198.18.0.0/15`, `198.51.100.0/24`,
  `203.0.113.0/24`, `224.0.0.0/4` (multicast), `240.0.0.0/4` and Azure's `168.63.129.16`.
- IPv6: anything outside global unicast `2000::/3`, which leaves out `::1`, `::`, IPv4-mapped and
  IPv4-compatible forms, NAT64 `64:ff9b::/96`, unique local `fc00::/7` (AWS metadata at
  `fd00:ec2::254`), link-local `fe80::/10` and multicast `ff00::/8`; and inside it `2001::/23`
  (Teredo among it), `2001:db8::/32`, `2002::/16` (6to4) and `3fff::/20`. A scoped address
  (`fe80::1%eth0`) is refused too.

The connection then goes to the address that was checked, with the host name kept for TLS (SNI
and the certificate check) and the `Host` header. Nothing looks the name up a second time, so a
DNS answer that changes between the check and the connection cannot point the request inside
the network. Redirects are never followed (a 3xx answer is a failure), the whole exchange has a
10 second deadline, at most 64 KB of the answer is read and none of it is kept, and no proxy from
the environment is used. Certificates are always verified.

The address carries the webhook's own token (Teams and Slack put it in the path or query), so
Prism365 treats it as a secret. It is stored only encrypted, with the state store's cipher
(AES-256-GCM under `STATE_ENCRYPTION_KEY`), bound to its tenant and destination so a stored value
cannot be moved to another row. After saving, the page shows the host only, and no log line or
ledger entry holds the address. Without `STATE_ENCRYPTION_KEY` the key lives in memory for the
life of the process, so saved webhooks stop working at a restart; the page says so, and such a
destination shows as unreadable until it is added again.

### Signatures

Each webhook gets its own signing secret when it is created. It is shown once, in the answer to
that request, and stored encrypted next to the address. Every request carries:

| Header | Value |
|---|---|
| `x-prism365-delivery` | The delivery id, as the delivery log shows it. |
| `x-prism365-timestamp` | Seconds since 1970, UTC. |
| `x-prism365-signature` | `v1=` and the hex HMAC-SHA256, keyed with the secret, over the timestamp, a dot and the raw body. |

A receiver recomputes the signature over the raw body, compares in constant time, and refuses a
timestamp more than five minutes away:

```js
import { createHmac, timingSafeEqual } from "node:crypto";

function fromPrism365(secret, headers, rawBody, now = Date.now()) {
  const timestamp = headers["x-prism365-timestamp"] ?? "";
  if (!/^[0-9]{1,12}$/.test(timestamp) || Math.abs(now / 1000 - Number(timestamp)) > 300)
    return false;
  const expected = Buffer.from(
    `v1=${createHmac("sha256", secret).update(`${timestamp}.${rawBody}`).digest("hex")}`,
  );
  const given = Buffer.from(headers["x-prism365-signature"] ?? "");
  return expected.length === given.length && timingSafeEqual(expected, given);
}
```

Teams and Slack ignore these headers; the generic format is where they matter.

### The generic format, version 1

```json
{
  "version": 1,
  "type": "prism365.alert",
  "deliveryId": "6f7b2c1e-0d3a-4c55-9a51-2f0c8e1b7d44",
  "sentAt": "2026-10-24T12:00:00.000Z",
  "kind": "finding",
  "title": "2 new or escalated findings, the worst at critical severity",
  "severity": "critical",
  "count": 2,
  "lines": ["Rule: High findings to the security channel"],
  "items": [
    {
      "title": "Sign-in attempt on a canary account",
      "severity": "critical",
      "subject": "Training Account",
      "note": "new",
      "details": {}
    }
  ],
  "link": "https://prism365.example/#/signins"
}
```

`kind` is `finding`, `directory`, `collection`, `audit`, `export`, `summary` or `test`.
`severity` is null for alerts without one. `count` is how many things the alert covers; `items`
lists at most ten of them. `note` is `new`, `reopened` or `escalated`. `details` is empty unless
the destination includes details. A later version changes `version`; fields are added, never
renamed, within one.

## Rules

A rule says what to watch and which destinations (one to five) to tell. A tenant can have 50.
Rules are the setting `alerts.rules`, kept exactly as stored: a save changes only what it is
about (a rule added, one rule turned on or off, one removed) and writes every other rule and field
back untouched. A rule this version cannot read, or one with a field it does not know (a newer
version may have added a field that narrows it), is not run at all and is listed on the page as
kept but not run, so a damaged or newer setting can only stop an alert, never widen one. Such rules
count towards the 50 and can be deleted. Scheduled exports are kept the same way. Four kinds:

- Findings: a new, reopened or escalated finding at or above a severity, optionally only from
  some modules.
- Directory changes: a directory change finding of the chosen kinds (a privileged role assigned,
  an app given a secret, consent granted, a Conditional Access policy changed, federation changed,
  sign-in methods removed by an administrator).
- Collection: collection that is meant to be on (an administrator started it and nobody stopped
  it) but has been stopped for at least the chosen hours, or a module reporting a span it could
  not read (a gap finding), or both.
- An audit number: one of an audit's tiles going above or below a value, or moving by at least a
  value, against its previous comparable run. Comparable means what the trend chart means: the
  same audit version, the same scope, and complete. A partial run never fires a rule. Above and
  below fire when the number crosses the line, not on every run past it.

Each rule has a quiet period, 5 minutes to 7 days: at most one alert per rule and subject in it.
The subject is the finding for finding, directory change and gap alerts, the tile for an audit
rule, and collection itself for stopped collection, which therefore reminds once per quiet period
while it stays stopped. A delivery that failed still counts, so a broken destination is
not tried again for the same thing until the period ends.

Findings are taken from the collection scheduler's sink, gathered for two seconds and sent as one
alert per rule and destination, worst first. The sink takes in only findings an enabled rule
wants. It holds 1,000; when it is full, the lowest-severity, oldest finding makes room for a new
one (or the new one is left out when it is the lowest), so a burst never pushes out a critical
finding. What is left out shows on the page as a failed delivery ("dropped") with its count, for
each rule that wanted it, and in the ledger as `alerts.dropped`. Audit rules are checked when a run's result is
stored. Collection is checked every five minutes. No rule reads Microsoft.

## What an alert holds

By default an alert holds its title, severity, counts, what it is about (the finding's subject
label) and a link back to Prism365 on `PUBLIC_ORIGIN`. With "include details" on, it also holds
each finding's summary and the plain fields of its first evidence record, at most twelve.

A full IP address leaves the server only when the destination is email (whose recipients are
always in `ALERT_EMAIL_DOMAINS`) and its details are on. Everywhere else every address in every
string, the subject label and the summary included, is cut to its network: `/24` for IPv4,
`/48` for IPv6. Percent-encoded colons and dots are read as what they stand for, and a run of
digits and dots or hex groups that is not one address but holds one (a port, a digit or a dotted
part glued on) is cut whole, so a version number with four or more parts is cut too. A failed sign-in pattern's subject is its address, so this matters even without
details.

Provider text is untrusted. For Teams and Slack, square brackets lose their link meaning and URL
schemes are broken (`https[:]//`), so nothing in it becomes a one-click link; for Slack, `&`, `<`
and `>` are escaped and `@channel`, `@here` and `@everyone` are broken, so it cannot mention
anyone. The one link in a message is Prism365's own. Every line is one line.

## Delivery

Delivery runs in the API process, off every request path. The scheduler's sink only adds a
finding to a buffer and returns, so a destination that hangs cannot hold up collection or audits.
Sending happens on three workers, and a destination never has more than one attempt in progress,
so one that hangs holds up only itself. An attempt that may work later (a timeout, a connection
or DNS failure, a 5xx answer, a 429) is tried again after 30 seconds and then 2 minutes, three
attempts in all; then the delivery is marked failed. A refusal (blocked address, 4xx, a redirect,
a TLS failure, a recipient outside the domains) is not tried again. An attempt that reaches its
time limit is aborted and its connection closed, so nothing keeps running beside the next one.
Email that loses its connection after the relay asked for the message is `outcome_unknown`: it
may have arrived, so it is never sent again.

Every delivery is one row in the delivery log, `prism365_delivery`: the destination, what sent
it (a rule, a test, an export or the summary), how many items it covered, its status, the number
of attempts and the last error as a fixed category. Never the message. The page shows the
newest 50. Rows are kept 90 days. The message itself is held in memory only until it is sent; a
delivery a restart interrupted is marked so when the API starts again.

Send test sends a fixed message once, without retries, and shows how it went.

## Scheduled exports

An audit set to run on its own can send each scheduled run to a destination. Email with details
on gets the run's CSV attached, as the export button writes it, with the same name (the scope
suffix included) and up to 5 MB. Every other destination, and a larger file, gets the counts and
a link back. The file carries every column, names and sign-in names included, which is why it
needs details on. A run an administrator started is not sent.

## Weekly summary

One destination can get the This week page as counts every Monday at 07:00 in
`SIGNINS_TIME_ZONE` (UTC when unset): new sign-in findings by severity, directory changes,
accounts that went quiet, audit numbers that moved, Microsoft deadlines and incidents, and how
collection went. No names, addresses or titles. A server that was down at 07:00 sends it when it
is back the same day; a week it missed entirely is skipped.

## Access ledger

Each change is recorded before it is made, and each delivery before anything leaves. Entries
hold ids and counts, never an address, a recipient, a name or the message.

| Entry | Fields |
|---|---|
| `alerts.destination.created` | destination id, kind, number of recipients, details on or off |
| `alerts.destination.deleted` | destination id |
| `alerts.rule.created` | rule id, kind, number of destinations |
| `alerts.rule.changed` | rule id, on or off |
| `alerts.rule.deleted` | rule id |
| `alerts.export.changed` | audit id, destination id or null |
| `alerts.weekly.changed` | destination id or null |
| `alerts.delivery` | source, rule id, destination id, delivery id, items; an export adds the audit and run ids, the summary the sections it read |
| `alerts.dropped` | rule id, how many findings a full buffer left out |

A rule's deliveries are recorded under the administrator who created the rule, an export's and
the summary's under the one who chose their destination, and a test under the one who sent it.

## Routes

| Route | Does |
|---|---|
| `GET /v1/alerts` | Settings, destinations, rules, exports, the weekly summary and recent deliveries |
| `POST /v1/alerts/destinations` | Adds a destination; a webhook's answer holds its signing secret, once |
| `DELETE /v1/alerts/destinations/:id` | Deletes one nothing sends to |
| `POST /v1/alerts/destinations/:id/test` | Sends a test and answers with the delivery |
| `POST /v1/alerts/rules` | Adds a rule, on |
| `POST /v1/alerts/rules/:id` | Turns a rule on or off (`{"enabled": false}`) |
| `DELETE /v1/alerts/rules/:id` | Deletes a rule |
| `POST /v1/alerts/exports` | `{"auditId", "destinationId"}`, null to stop |
| `POST /v1/alerts/weekly` | `{"destinationId"}`, null for none |

Every route needs a signed-in administrator; the others also need the CSRF token, as every
change does.

## Not done

No Graph mail, SMS, push or inbound webhooks, no per-user alert preferences, and no remediation.
