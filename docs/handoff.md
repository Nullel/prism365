# Handoffs and receipts

Status: built (TL-033). Code: `packages/contracts/src/handoff.ts` (both formats and the handoff
builder), `apps/api/src/routes/handoffs.ts` (the routes), `apps/api/src/handoff-store.ts` and core
migration `006-handoffs.sql` (storage), `apps/web/src/Handoffs.tsx` and
`apps/web/src/PersonReceipts.tsx` (the screens).

Prism365 sees; it never changes a tenant. Acting on what it finds, such as blocking sign-in for
an account nobody has used in a year, belongs to a separate action tool: an offboarding script, a
lifecycle product, whatever the organisation uses. The two meet through files. Prism365 exports a
handoff, the list of accounts to act on with the reasons and evidence. The action tool does its
work and writes a receipt saying what it did to whom. Prism365 imports the receipt, keeps it, and
shows it on the Handoffs page and on each person's page.

Prism365 never calls the action tool, never receives its credentials, and never gives it any of
its own. The only thing that crosses is the two files, carried by the administrator.

## 1. Exporting a handoff

On an audit's results, **Export for action** sits next to Export CSV. It is offered only by the
audits whose rows are accounts an action tool could act on, which say so in their definition
(`handoff: true`): Inactive accounts, Licenses to reclaim, MFA registration, Guest accounts and
Risky users. A person column alone is not enough: Managed devices names each device's owner, but
its rows are devices. Microsoft 365 app usage names accounts only by sign-in name, so its rows
carry no directory id to act on. App consents, Conditional Access policies and License capacity
have no handoff either.

An account is handed off by its directory object id, so a run whose rows carry none has nothing
to hand off and is refused, and the button is not shown for it.

The export holds exactly the accounts the page shows: the same view, filters, search and sort as
the CSV export, read by the same code on the server, so the CSV and the handoff of one view
always hold the same accounts in the same order. Rows that are not an account are left out. One
difference: the CSV export takes a view the audit does not have to mean every row, and the
handoff refuses it (400), so it never hands off accounts nobody chose.

The first time a browser exports one, a note under the toolbar says what a handoff is: a list for
the tool that acts on it, and Prism365 itself changes nothing.

The file is named `prism365-handoff-<audit>-<yyyy-mm-dd>.json`, with the day of the run, as the
CSV is. A run limited to some account domains says so in the name as the CSV does
(`prism365-handoff-accounts.inactive-2026-10-22-limited-to-2-domains.json`), and in the file.

## 2. The handoff format, version 1

A JSON object. Every object in it is strict: a reader should refuse a field it does not know,
because that means a later version.

| Field | Meaning |
|---|---|
| `format` | Always `"prism365.handoff"` |
| `version` | Always `1` |
| `id` | A UUID Prism365 chose for this export. Put it in the receipt's `handoffId`. |
| `generatedAt` | When it was exported (ISO 8601, UTC) |
| `tenantId` | The Microsoft Entra tenant the accounts belong to |
| `generatedBy` | The administrator who exported it: `actorId` (their directory object id) and `name` (or null) |
| `source.kind` | Always `"audit"` |
| `source.auditId`, `source.auditTitle`, `source.auditVersion` | The audit, its title, and the version of it that produced the rows |
| `source.runId`, `source.runFinishedAt` | The run the rows come from, and when it finished |
| `source.scope` | The account domains the run was limited to, such as `["larkspur.example"]`, or null when it covered every domain. A limited list is not the whole tenant. |
| `source.complete` | false when Microsoft returned less than everything to the run (a record it could not read, a read limit), so the list may miss accounts |
| `source.view` | The view the rows were taken from, `{ "id", "label" }` |
| `source.search` | The search applied, or `""` |
| `source.facets` | The filters applied, by column key, such as `{ "status": "Disabled" }` |
| `source.sort` | `{ "column", "direction" }`, or null for the audit's own order |
| `items` | The accounts, at most 20,000, in the order of the page |
| `items[].userId` | The account's directory object id. Act on this. |
| `items[].userPrincipalName`, `items[].displayName` | As the run read them, or null |
| `items[].reasons` | Why the account is in the list: the labels of the audit's views it falls in, such as `"90+ days"`, in the audit's order. At most 20. |
| `items[].evidence` | The row's values in the columns the page showed, by column label: a string (at most 2,000 characters), a number, true or false, or null for no value. Lists are joined with `; `, dates are ISO 8601. At most 40. |

A complete example, from the inactive accounts audit's 90+ days view:

```json
{
  "format": "prism365.handoff",
  "version": 1,
  "id": "5d0f8a2c-7b14-4e39-a6c1-0f2e9d8b7a63",
  "generatedAt": "2026-10-22T14:05:11.000Z",
  "tenantId": "0a1b2c3d-4e5f-4a6b-8c7d-9e0f1a2b3c4d",
  "generatedBy": {
    "actorId": "6e5d4c3b-2a19-4f08-9e7d-6c5b4a392817",
    "name": "Morgan Example"
  },
  "source": {
    "kind": "audit",
    "auditId": "accounts.inactive",
    "auditTitle": "Inactive accounts",
    "auditVersion": 3,
    "runId": "c4b3a291-8f7e-4d6c-9b5a-493827160514",
    "runFinishedAt": "2026-10-22T13:59:02.000Z",
    "scope": null,
    "complete": true,
    "view": {
      "id": "d90",
      "label": "90+ days"
    },
    "search": "",
    "facets": {},
    "sort": {
      "column": "days",
      "direction": "desc"
    }
  },
  "items": [
    {
      "userId": "9b2e7d14-5a63-4f08-b1c9-2e8f7a6d5c40",
      "userPrincipalName": "blake@larkspur.example",
      "displayName": "Blake Example",
      "reasons": [
        "7+ days",
        "30+ days",
        "60+ days",
        "90+ days",
        "120+ days",
        "180+ days",
        "1 year+",
        "Never signed in"
      ],
      "evidence": {
        "Inactive for": 412,
        "Counted from": "Account created",
        "Last successful sign-in": null,
        "Status": "Enabled",
        "Type": "Member",
        "Licenses": null,
        "Created": "2025-09-05T09:30:00.000Z"
      }
    },
    {
      "userId": "3f6c2b9e-8d41-4c7a-9e25-6b1d0a7f4c38",
      "userPrincipalName": "avery@larkspur.example",
      "displayName": "Avery Example",
      "reasons": [
        "7+ days",
        "30+ days",
        "60+ days",
        "90+ days"
      ],
      "evidence": {
        "Inactive for": 97,
        "Counted from": "Successful sign-in",
        "Last successful sign-in": "2026-07-17T08:12:00.000Z",
        "Status": "Enabled",
        "Type": "Member",
        "Licenses": "Microsoft 365 E3",
        "Created": "2024-03-04T10:00:00.000Z"
      }
    }
  ]
}
```

## 3. Using a handoff in an action tool

1. Check `format` and `version`, and refuse anything else.
2. Check `tenantId` is the tenant the tool is about to act on.
3. Show the operator the list, `source.auditTitle`, `source.view.label` and `source.runFinishedAt`:
   the list is only as current as the run. Say so when `source.scope` is not null or
   `source.complete` is false.
4. Act on each item by `userId`. A principal name can be renamed or reused; the object id cannot.
   Read the account again before acting: the handoff says what Prism365 saw, not what is true now.
5. Write one receipt for the run, with a new UUID as its `id` and the handoff's `id` as its
   `handoffId`. Report every action you took or decided not to take, one item each.

The handoff holds names, sign-in names and account details. Treat the file like the CSV export:
it is outside Prism365's control once downloaded.

## 4. The receipt format, version 1

A JSON object, strict at every level like the handoff: an unknown field anywhere is refused.

| Field | Meaning |
|---|---|
| `format` | Always `"prism365.receipt"` |
| `version` | Always `1` |
| `id` | A UUID the tool chooses for this receipt. Importing it twice keeps it once (section 5). |
| `handoffId` | The `id` of the handoff this answers, or null for actions taken without one. It must be a handoff exported from the same Prism365 for the same tenant. |
| `tool.name`, `tool.version` | The tool, 1 to 80 characters, and its version, 1 to 40. The name cannot be Prism365 (in any letter case or width): a receipt reports what another tool did. |
| `performedAt` | When the tool did the work, ISO 8601 with a time zone (`Z` or an offset of at most 15:59), between the years 2000 and 2099 and no more than a day ahead of Prism365's clock. Kept in UTC. |
| `performedBy` | The operator, as the tool names them, 1 to 200 characters |
| `items` | What was done, 1 to 20,000 items |
| `items[].userId` | The account's directory object id, or null |
| `items[].userPrincipalName` | The account's principal name, at most 500 characters, or null. At least one of the two is required. |
| `items[].action` | What was done, as a lower-case name of 2 to 80 letters, digits, dots and dashes starting with a letter, such as `account.sign-in-blocked`, `license.removed` or `mailbox.converted-to-shared`. Names starting with `prism365.` are kept for Prism365. Prism365 shows it as words ("Account sign in blocked"). |
| `items[].status` | `done`, `failed` or `skipped` |
| `items[].detail` | Anything the operator should know, 0 to 500 characters |

Before text is kept, it is trimmed (except `detail`), NUL characters are removed, broken UTF-16
is replaced with U+FFFD, and characters that reorder how text is displayed (U+202A to U+202E,
U+2066 to U+2069, U+200E and U+200F) are removed. Text over its limit is refused, never cut. Ids
that are GUIDs are kept in lower case.

A complete example, answering the handoff above:

```json
{
  "format": "prism365.receipt",
  "version": 1,
  "id": "e8d7c6b5-a493-4827-9160-5f4e3d2c1b0a",
  "handoffId": "5d0f8a2c-7b14-4e39-a6c1-0f2e9d8b7a63",
  "tool": { "name": "Example action tool", "version": "1.4.0" },
  "performedAt": "2026-10-23T09:30:00Z",
  "performedBy": "Jordan Vale",
  "items": [
    {
      "userId": "9b2e7d14-5a63-4f08-b1c9-2e8f7a6d5c40",
      "userPrincipalName": "blake@larkspur.example",
      "action": "account.sign-in-blocked",
      "status": "done",
      "detail": "Never used since it was created. Sign-in blocked after review."
    },
    {
      "userId": "3f6c2b9e-8d41-4c7a-9e25-6b1d0a7f4c38",
      "userPrincipalName": "avery@larkspur.example",
      "action": "account.sign-in-blocked",
      "status": "skipped",
      "detail": "On leave until November; the manager asked to keep it."
    },
    {
      "userId": "3f6c2b9e-8d41-4c7a-9e25-6b1d0a7f4c38",
      "userPrincipalName": null,
      "action": "license.removed",
      "status": "failed",
      "detail": "The tool's account lacks the License Administrator role."
    }
  ]
}
```

## 5. Importing a receipt

On the Handoffs page (Tools in the navigation), **Import receipt** reads the chosen file in the
browser and sends it. The page says how many actions were done, failed and skipped, or what is
wrong with the file, naming the first problem and where it is (`items[3].action is not an action
name ...`).

- Importing is idempotent by id. A receipt's digest is the SHA-256 of its body as received, in
  canonical JSON (keys sorted, no spaces), taken before any validation or cleaning, so sending
  the same file again always gives the same digest, and key order and spacing never change it.
  Any other difference does, including the same time written another way (`Z` or `+00:00`).
  Sending a receipt whose id is already kept answers 200 with what was kept when the digest is
  the same, and 409 when it is not. Each digest is stored with the version of the method that
  took it (`digest_version`, now 1); a receipt kept under another version cannot be compared and
  is answered 200 as kept. A kept receipt never changes: a tool corrects itself with a new
  receipt under a new id.
- A `handoffId` that is not a handoff of this tenant is refused with 400.
- The receipt and its items are stored in one transaction: if any item fails to store, nothing
  of the receipt is kept, and sending it again imports it.
- A receipt is at most 4 MB of JSON with at most 20,000 items. A tool with more splits its work
  into several receipts, each with its own id. A request that is not from a signed-in
  administrator is refused before its body is read.

There is no way to edit or delete a receipt in Prism365, and no route for it. Receipts are
deleted only by retention (section 8).

## 6. What receipts are worth

A receipt is a claim made by another tool. Prism365 stores it and shows it as reported, labelled
as that tool's report, and never treats it as checked state: it does not mark an account blocked
because a receipt says so, and it does not compare receipts with audit results. To see the
account as it is, open the person's page or run the audit again.

## 7. On the person page

**Reported by other tools** lists what receipts say was done to the account, newest first (by
`performedAt`), at most 100, and says when there are more: the action, the status, the tool and
its operator, when, and the detail. An item matches the account by its `userId`, or, when the
tool gave only a principal name, by that name. Names are compared by a key Prism365 makes itself,
NFKC-normalised and in lower case, so letter case and look-alike forms (the Kelvin sign for K,
fullwidth letters) never decide a match; the name is still shown as the tool wrote it. Items
matched by name are marked "Matched by sign-in name", because a name can move to another
account. An item that names a different account's id never matches by name. The section is left
out when nothing matches.

## 8. Storage, retention and the access ledger

Core migration `006-handoffs.sql` adds three tenant tables with forced row-level security:

- `prism365_handoff`: one row per export: the audit, run, view, whether a search or filter
  narrowed it, the run's domain scope and whether it was complete, how many accounts, when and by
  whom. The accounts themselves are only in the file.
- `prism365_receipt`: one row per receipt as sent: the tool, operator, when, counts by status,
  when and by whom it was imported, and its digest and digest version. Keyed by tenant and id,
  so one tenant's receipt ids never collide with another's.
- `prism365_receipt_item`: each reported action, in order, indexed by account id and by the
  principal name key.

Handoffs and receipts are kept 3 years by default. `HANDOFF_RETENTION_DAYS` (30 to 3650) changes
that. Whenever a new handoff or receipt is stored, in the same transaction, Prism365 deletes the
receipts imported before the retention period with their items, and the handoffs exported before
it that no kept receipt answers. Nothing else deletes them. See also `docs/PRIVACY.md`.

The access ledger gets one entry for each export, import and read, with ids and counts, never
names or anything the tool wrote:

- `handoff.exported` (audit, run, handoff, number of accounts). The export reads the stored run
  first, then writes the entry, then stores the handoff row and sends the file.
- `receipt.imported` (receipt, handoff, number of items, and how many were done, failed and
  skipped). The import validates the receipt and reads whether its id is already kept and its
  handoff exists, then writes the entry, then stores the receipt. An invalid receipt and a
  repeat of a kept one store nothing and write no entry. An import that loses a race for its id
  to another at the same moment has already written its entry when the insert finds the winner,
  so the ledger can hold an entry for an import that stored nothing, never the other way round.
- `handoffs.viewed`, `receipt.viewed` (receipt) and `person.receipts.viewed`: written before the
  read, so nothing is read without one.

## 9. API

All routes need a signed-in administrator; the POST also needs the CSRF header.

| Route | Returns |
|---|---|
| `GET /v1/audit-runs/{runId}/handoff.json` | The handoff file. Takes the CSV export's `view`, `search`, `sort`, `direction` and `f.<column>`, plus `columns` (comma-separated keys of the columns to use as evidence; the audit's default columns without it). 404 for an unknown or unfinished run, an audit without a handoff, and a run whose rows name no account by id; 400 for a view the audit does not have. |
| `POST /v1/receipts` | 201 `{ receipt, imported: true }` for a new receipt; 200 `{ receipt, imported: false }` for one already kept; 409 for a different receipt under a kept id; 400 for an invalid receipt or an unknown handoff; 413 over 4 MB |
| `GET /v1/handoffs` | The newest 100 handoffs and the newest 100 receipts, each with counts by status and the handoff it answers |
| `GET /v1/receipts/{id}` | One receipt with its items in order; `limit` (1 to 500, default 200) and `after` (the `next` of the previous page) |
| `GET /v1/people/{userId}/receipts` | The items that name this account, newest first, at most 100, with `truncated` when there are more; `principal` adds items that name only that principal name |
