# ADR-0012: shared mailbox access comes from a read-only export the administrator runs

- Status: accepted
- Date: 2026-11-06
- Owners: maintainers

## Context

Who holds Full Access, Send As and Send on Behalf on a shared mailbox is one of the first
questions in an offboarding or a compromise review, and the person page's Access tab could not
answer it. We checked every documented read path on 2026-11-06:

- Microsoft Graph exposes none of the three rights.
- The Exchange Online Admin API (preview) has no permission cmdlets, and needs `Exchange.ManageV2`,
  a scope that can manage, not only read.
- Exchange Online PowerShell has read-only cmdlets for all three: `Get-EXOMailboxPermission`
  (Full Access), `Get-EXORecipientPermission -AccessRights SendAs` (Send As) and the mailbox's
  `GrantSendOnBehalfTo` from `Get-EXOMailbox -Properties GrantSendOnBehalfTo`. Running them needs an
  Exchange session and an Exchange role, which Prism365 does not have and should not hold: an
  Exchange credential on the server would be a write-capable credential in a read-only product.

## Decision

An administrator runs `Export-MailboxAccess.ps1` (served by the web app at
`/scripts/Export-MailboxAccess.ps1`) in their own `Connect-ExchangeOnline` session and imports the
JSON file it writes on the Microsoft access page. Prism365 never calls Exchange, never holds an
Exchange credential and asks Microsoft for no new permission.

- The script calls only Get- cmdlets and writes only its output file; a test reads it and fails
  on anything else. It refuses to run with no Exchange Online connection or with more than one,
  and takes the tenant from `Get-ConnectionInformation`. It resolves each trustee once with
  `Get-EXORecipient` and never with an empty value (Exchange then answers with every recipient);
  a trustee that does not resolve to exactly one recipient is written as `unresolved` with the
  text Exchange returned; any other lookup error stops it. It stops, writing nothing, when Exchange
  refuses a cmdlet, and names the role check `Get-ManagementRole -Cmdlet Get-MailboxPermission`.
- Prism365 serves code that runs with an Exchange admin session, so the script's SHA-256 is
  published in `docs/SETUP.md`, the page shows the hash of the copy it serves and whether it is the
  one this version shipped, and both tell the administrator to read the script and check it with
  `Get-FileHash` against the published hash before running it.
- The import (`POST /v1/mailbox-access/import`, up to 25 MB on that route only) validates the whole
  file before anything is recorded or stored: the format and version, this installation's tenant,
  an export date neither in the future nor older than 90 days, at most 20,000 mailboxes and 200,000
  grants, GUID ids, bounded and cleaned text, and a file that agrees with itself (every grant's
  mailbox listed, no repeated grant, the unresolved count right). The ledger entry comes first.
- Only the latest import is kept (core migration `008-mailbox-access.sql`). A new import replaces
  the previous one in one transaction, so a failure keeps the old one whole; a clear deletes it.
  Both first lock a per-tenant row that always exists once created, so they take turns even
  before the first import exists.
- The person page shows the rights the import gives the person directly and through groups, which
  are matched by the person's groups, nested ones included, read live with `directory.users.read`
  (`transitiveMemberOf`). Without that pack only direct grants are matched, and the page says so.
  When the person is a mailbox in the import, the page lists who holds which right on it.

## Consequences

The data is a snapshot that the administrator decides when to refresh, not something Prism365
watches; the pages always show the export's date and warn when it is older than 30 days. Anyone
who can edit the file can make it claim anything, so the file is treated as untrusted input and
its contents as the export's claim: the pages show provenance, Prism365 never acts on it, and
nothing else in Prism365 (audits, alerts, handoffs) reads it. Which Exchange role may run the
permission cmdlets is not documented; Global Reader maps to View-Only Organization Management, and
the setup guide tells the administrator to check with `Get-ManagementRole`. Mailbox folder
permissions and scheduled exports are out of scope.
