# ADR-0009: mail forwarding waits for the application credential

- Status: accepted
- Date: 2026-10-23
- Owners: maintainers

## Context

Mail that leaves the organisation through forwarding is one of the first things an attacker sets
up after taking over a mailbox, so "which mailboxes forward mail, and where" belongs in the eyes.
Prism365 reads with the delegated session of a signed-in administrator, and only with GET
requests (ADR-0002, `ROADMAP.md` Milestone 6). We checked every documented read path against that
model on 2026-10-23:

- Microsoft Graph does not expose mailbox forwarding at all. `mailboxSettings` has no
  `ForwardingSmtpAddress`, `ForwardingAddress` or `DeliverToMailboxAndForward`.
- Inbox rules that forward or redirect are readable at
  `/users/{id}/mailFolders/inbox/messageRules` with `MailboxSettings.Read`. With a delegated
  token that reaches only mailboxes the signed-in administrator has been delegated, not every
  mailbox. Reading every mailbox needs the application permission `MailboxSettings.Read`.
- The Exchange Online Admin API (preview) is POST-only even for reads, needs `Exchange.ManageV2`
  and the Recipient Management role group (both write-capable), exposes no inbox rules, and says
  its `Get-Mailbox` output will be cut to core properties without forwarding before release.
- `Get-InboxRule` in Exchange Online PowerShell does not work for the read-only roles (View-Only
  Organization Management, Global Reader).
- Purview audit search shows forwarding changes (`Set-Mailbox`, `New-InboxRule`) inside the
  retention window, but starts every search with a POST, and shows changes, not current state.

## Decision

Mail forwarding is not built on the delegated session. It becomes part of the application
credential milestone (ADR-0002, sign-in audit S5): when Prism365 gains an app-only identity, the
forwarding module reads inbox rules for every mailbox with the application permission
`MailboxSettings.Read`, limited to the mailboxes the operator chooses with RBAC for Applications.
Admin-set SMTP forwarding stays out of scope until Microsoft offers a read-only way to see it.

## Consequences

The forwarding monitor's Graph half is not ported now. The directory change feed does not cover
forwarding either, because mailbox settings are Exchange events, not directory audits. The
outbound spam policy (which can block automatic external forwarding for the whole tenant) is the
practical control meanwhile; an administrator checks it in the Defender portal. When the
application credential lands, this decision is revisited with its own threat model, because an
app-only permission works without a signed-in administrator.
