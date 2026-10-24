# ADR-0011: alerts go through the operator's relay or a checked webhook, never Graph mail

- Status: accepted
- Date: 2026-10-24
- Owners: maintainers

## Context

Prism365 finds things an administrator should hear about without opening it: a canary account
signing in, a privileged role assigned, collection that stopped. Telling them means sending
something out of the server, and Prism365's first promise is that it never writes to Microsoft.

Three ways to send were open:

- Graph mail (`Mail.Send`). It needs a write permission on mailboxes, which breaks the read-only
  promise in `docs/PERMISSIONS.md` and would let a compromised Prism365 send mail as someone in
  the tenant. The Graph client refuses every method but GET, and that stays true.
- The operator's own SMTP relay. The operator already controls it, its credentials stay outside
  the tenant, and it can only send.
- Webhooks to Teams Workflows, Slack or another endpoint. These are what most teams watch, but an
  address an administrator types is still an address the server connects to: pointed at
  `169.254.169.254`, `localhost` or the internal network, it is a server-side request forgery.
  Teams and Slack addresses also carry their own token, so the address is a secret.

What leaves the server matters as much as how. A finding's evidence can hold names, sign-in
names and full IP addresses, and a chat channel is read by more people than the Prism365
administrators.

## Decision

Alerts are sent through the operator's mail relay or a webhook, and never through Microsoft.

- Email is configured by environment only: `SMTP_URL` (smtps, or smtp with STARTTLS required),
  the password from `SMTP_PASSWORD_FILE` (never from the environment, the page or a request),
  `SMTP_FROM`, and `ALERT_EMAIL_DOMAINS`, which every recipient must be in exactly, checked when
  saved and before each message. Certificates are verified. Header values never carry a line
  break, and recipients are passed as mailboxes and envelope, never parsed from a string.
- A webhook address must be HTTPS with a host name, no user info, and port 443 unless
  `WEBHOOK_ALLOWED_PORTS` lists another. The host is resolved at send time; if any answer is
  loopback, private, link-local, CGNAT, multicast, reserved, documentation, a cloud metadata
  address, or an IPv6 form that carries IPv4 (mapped, compatible, NAT64, 6to4, Teredo), nothing is
  sent. The connection goes to the checked address, with the host name for TLS and `Host`, so a
  second lookup can never be swapped in. No redirects, a 10 second deadline, at most 64 KB read,
  no environment proxy.
- A webhook's address and its server-generated signing secret are stored only encrypted with the
  state store's cipher, bound to tenant and destination. The page shows the host afterwards; the
  secret is shown once. Requests are signed with HMAC-SHA256 over timestamp and body.
- By default an alert carries its title, severity, counts, subject label and a link back. A
  destination's "include details" adds evidence fields. A full IP address leaves only for email
  with details on; everywhere else addresses are cut to `/24` or `/48`. Chat formats neutralise
  links and mentions in provider text.
- Delivery runs in the API process off every request path and never blocks collection or audits.
  Each delivery is one log row (status, attempts, error category, never the message) and one
  access-ledger entry (ids and counts), written before anything is sent. Three attempts with
  backoff, then failed.

## Consequences

- No new Microsoft permission. `docs/PERMISSIONS.md` keeps every `*.Send` and `*.ReadWrite.*`
  permission out.
- Operators who want email must run or rent a relay and give Prism365 its own account on it.
- A webhook that resolves to a private address cannot be used, even when the operator means it
  (an internal chat server). That is deliberate; an outbound proxy or a public endpoint is the
  way around it, and either is a choice the operator makes outside Prism365.
- Without `STATE_ENCRYPTION_KEY` webhook addresses are encrypted with a key held in memory and
  stop working at a restart.
- The message is never stored, so a delivery a restart interrupts is lost and marked so.
- A chat alert without details may name a person (the finding's subject label). Operators who do
  not want names in a channel should send those rules to email.

## Validation

`apps/api/test/delivery-address.test.ts` and `delivery-webhook.test.ts` cover every refused range
at its edges, the IPv6 forms, a DNS answer that changes between checks, redirects, user info,
ports and the deadline. `delivery-smtp.test.ts` covers TLS, the password file and the recipient
domains, `delivery-payload.test.ts` what leaves the server, `delivery-service.test.ts` and
`delivery-api.test.ts` the rules, retries, ledger and the sink, and the storage isolation suite
the tables on PGlite and PostgreSQL. Each was checked against a deliberate break.

## Revisit when

- Prism365 runs as more than one API process: the quiet period and retries live in one process.
- Teams or Slack change their webhook formats, or a destination needs a format of its own.
- Someone needs a private webhook endpoint badly enough to justify an operator-set allowlist of
  internal hosts, which would need its own review.
