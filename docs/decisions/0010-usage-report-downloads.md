# ADR-0010: one redirect, for usage report downloads

- Status: accepted
- Date: 2026-10-23
- Owners: maintainers

## Context

The Graph client (`packages/graph/src/client.ts`) refuses every redirect. A client that follows
one can carry the request, and the bearer token with it, to a host nobody reviewed, so
`request()` and `getCollection()` send with `redirect: "manual"` and refuse whatever 3xx comes
back.

Microsoft's per-user usage reports work differently. In Graph v1.0, `GET /reports/<report>(period='D7')`
answers `302 Found` with a `Location` on `reports.office.com/data/download/`, and the file there is
CSV. Microsoft's documentation says these "preauthenticated download URLs are only valid for a
short period of time (a few minutes) and don't require an Authorization header". Only
`getM365AppUserDetail` has a documented JSON form in v1.0 (`$format=application/json`); for the
others JSON is beta only, and Prism365 uses v1.0. Usage audits built on the email, Teams,
OneDrive, SharePoint, mailbox and active-user reports need those files, so Prism365 either
follows this redirect or does without them.

Checked on Microsoft Learn 2026-10-23:
[getOffice365ActiveUserDetail](https://learn.microsoft.com/en-us/graph/api/reportroot-getoffice365activeuserdetail?view=graph-rest-1.0),
[getEmailActivityUserDetail](https://learn.microsoft.com/en-us/graph/api/reportroot-getemailactivityuserdetail?view=graph-rest-1.0),
[report authorization](https://learn.microsoft.com/en-us/graph/reportroot-authorization) and
[throttling limits](https://learn.microsoft.com/en-us/graph/throttling-limits). Microsoft shows
`reports.office.com` in every example but does not state it as a contract, and does not document
the file's encoding.

## Decision

The client gets one narrow exception, `GraphHttpClient.getReport(path, context, options)`, and
no other request follows a redirect.

- Only the `reports.usage.read` pack may call it; any other pack is refused (reason `report`).
  The pack's path allowlist is exactly eight reports, each by period only (`D7`, `D30`, `D90`,
  `D180`), with no `date=` form: `getM365AppUserDetail`, `getOffice365ActiveUserDetail`,
  `getEmailActivityUserDetail`, `getTeamsUserActivityUserDetail`, `getOneDriveActivityUserDetail`,
  `getOneDriveUsageAccountDetail`, `getSharePointActivityUserDetail` and `getMailboxUsageDetail`.
  `getReport` refuses any query option. That is enforced in `getReport` only: the allowlist checks
  paths, and `request()` still takes query options on these paths, because the Microsoft 365 app
  usage audit reads the JSON form of `getM365AppUserDetail` with `$format=application/json`.
- The Graph request passes every check `request()` applies (pack permission, cloud, GET only,
  host, path allowlist, the `$expand` refusal) and the same retry rules: 429 and 503 wait for
  `Retry-After` within the retry budget, other errors keep only their status and request ID.
  It is sent with the token and `redirect: "manual"`. Two answers are accepted: `302` with a
  `Location`, or `200` with a CSV body from Graph itself, in case Microsoft ever answers directly.
- The `Location` must parse, without a base, as an `https:` URL whose host is exactly
  `reports.office.com`, on the default port, with no user name or password and no fragment, and
  with a path under `/data/download/`. Anything else is refused before any download with a
  `GraphReportHostError` (reason `report_host`), which carries the host and a fixed problem, never
  the path or query. The audit page shows the host ("Microsoft sent the report from a host
  Prism365 does not recognise: ..."), and the server's run log line carries it with the
  `report_host` category, so an operator can report a host Microsoft has moved to without passing
  on the link. A parsed hostname is printable ASCII, cut to 253 characters.
- The token never goes to the download host. The download is a GET with no `Authorization`
  header and no cookies (`credentials: "omit"`), and Prism365 sets no header on it but `Accept`.
  Node's fetch adds its own non-secret headers to every request (on Node 26: `user-agent`,
  `accept-encoding`, `accept-language` and `sec-fetch-mode`, plus `Host`); none carries a
  credential. `redirect: "error"` means the download host cannot send the request on anywhere.
  It runs within the client's request timeout, the same budget as the Graph request, and the
  caller's signal, and is used once: nothing is retried.
- Any answer from the download host but 200, a partial (206) or empty (204) one included, is a
  `GraphReportDownloadError` with only its status, never the link or the host's text. It is not a
  `GraphRequestError`: a 401 or 403 there means the short-lived link failed, not that the session
  or a role is missing, so an audit says "Microsoft's report download link failed. Run the audit
  again." and a collector files it as a provider error. A failure inside fetch becomes a fixed
  error (reason `download`), shown the same way, because the underlying error could carry the link.
- The body must be `application/octet-stream` or `text/csv`, and is capped while it streams:
  16 MiB by default, at most 32 MiB (reason `size`). Fetch decompresses what it asked for with
  `accept-encoding`, and the cap counts the decompressed bytes, so a compressed body cannot slip
  past it.
- `packages/graph/src/csv.ts` parses it: a strict RFC 4180 reader that scans one character at a
  time, strips one leading byte-order mark, accepts CRLF or LF, requires a header row with
  unique, non-empty names, and allows at most 100 columns, 50,000 data rows (more sets
  `truncated`) and 8,192 characters a field. A row with a field over the limit, a field count
  other than the header's, or broken quoting is skipped and counted; an unterminated quote fails
  the report (reason `csv`). Values stay strings exactly as written. The result is the header,
  an `index` from header name to position, and each row as an array of strings, so callers find a
  column by name and a missing one reads as undefined. Arrays rather than one object per row keep
  a parsed report at about five times its size (measured 2026-10-23); objects took twelve.
- Modules and audits reach it through `ScopedGraph.getReport`, with the same declared-pack check
  and stop signal as their other reads.

## Consequences

- One reviewed path follows one redirect. `request()` and `getCollection()` still refuse every
  3xx, including for this pack, and tests hold both.
- No new permission. `Reports.Read.All` already covered these reports, and per-user detail still
  needs Reports Reader or a workload admin role. The `directory.users.read` pack also gains
  `/users/{id}/memberOf` (direct memberships, not transitive) for the audits that follow, since
  `User.Read.All` is Microsoft's documented least privilege for another user's memberships
  ([user-list-memberof](https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0)).
  Every key in that pack, for a user, its manager or its memberships, in the `/users/{id}` or the
  `users('upn')` form, now takes only what the path builders produce from an id or a principal
  name (what `encodeURIComponent` leaves as it is, including `! ~ * ' ( )`, the `:` and `@` of raw
  ids, and well-formed percent escapes), and no encoded slash, backslash or question mark. Microsoft
  allows none of those three in a principal name, so the builders answer such a name as not found
  without a request.
- A link is a credential for a few minutes. It is never stored, logged or put in an error, and is
  not cached.
- If Microsoft moves downloads to another host, reports fail with `report_host` until the list is
  changed in code and reviewed. That is deliberate: the list never comes from configuration or
  tenant input.
- Microsoft limits CSV reports to 14 requests per ten minutes per app per tenant, and does not say
  whether the download counts, so a caller should read each report once per run.
- Concealed names arrive as hashes in the same columns. The parser does not unmask anything; an
  audit that reads these reports handles concealment as the Microsoft 365 apps audit does.

## Validation

`packages/graph/test/report.test.ts` uses a fake fetch to check every request's headers (the token
goes to Graph only), each refused `Location` form before any download, a second redirect, the byte
cap mid-stream, timeouts and aborts on both requests, throttling, a direct CSV answer, the other
packs, and `request()` refusing a 302 for every pack. `packages/graph/test/user-paths.test.ts`
checks the user pack's keys against the path builders' own output, and
`apps/api/test/audit-failure.test.ts` what the audit page and the log say about a refused link or
a failed download. `packages/graph/test/csv.test.ts` covers the parser's rules and limits and a
10 MB report in well under a second. The synthetic tenant serves
one report through the same redirect (`scripts/demo/graph/usage-report-download.test.ts`) and
refuses a download that carries a token.

## Revisit when

- Microsoft documents a different download host, a JSON form in v1.0 for these reports, or a
  stated encoding.
- A report outgrows 32 MiB or 50,000 rows in a supported tenant.
- Any other feature asks to follow a redirect: it gets its own decision, not a wider list here.
