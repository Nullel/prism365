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
nothing else changes for any other request.

- Only the `reports.usage.read` pack may call it; any other pack is refused (reason `report`).
  The pack's path allowlist is exactly eight reports, each by period only (`D7`, `D30`, `D90`,
  `D180`), with no `date=` form and no query options: `getM365AppUserDetail`,
  `getOffice365ActiveUserDetail`, `getEmailActivityUserDetail`, `getTeamsUserActivityUserDetail`,
  `getOneDriveActivityUserDetail`, `getOneDriveUsageAccountDetail`,
  `getSharePointActivityUserDetail` and `getMailboxUsageDetail`.
- The Graph request passes every check `request()` applies (pack permission, cloud, GET only,
  host, path allowlist, the `$expand` refusal) and the same retry rules: 429 and 503 wait for
  `Retry-After` within the retry budget, other errors keep only their status and request ID.
  It is sent with the token and `redirect: "manual"`. Two answers are accepted: `302` with a
  `Location`, or `200` with a CSV body from Graph itself, in case Microsoft ever answers directly.
- The `Location` must parse, without a base, as an `https:` URL whose host is exactly
  `reports.office.com`, on the default port, with no user name or password and no fragment, and
  with a path under `/data/download/`. Anything else is refused before any download (reason
  `report_host`) with a fixed message that names the host, never the path or query, so an
  operator can report a host Microsoft has moved to without passing on the link.
- The token never goes to the download host. The download is a GET with no `Authorization`
  header, no cookies (`credentials: "omit"`), no header but `Accept`, and `redirect: "error"`, so
  the download host cannot send it on anywhere. It runs within the client's request timeout and
  the caller's signal, and is used once: a non-200 answer is a `GraphRequestError` with no
  provider text, and nothing is retried. A failure inside fetch becomes a fixed error (reason
  `download`), because the underlying error could carry the link.
- The body must be `application/octet-stream` or `text/csv`, and is capped while it streams:
  16 MiB by default, at most 32 MiB (reason `size`).
- `packages/graph/src/csv.ts` parses it: a strict RFC 4180 reader that scans one character at a
  time, strips one leading byte-order mark, accepts CRLF or LF, requires a header row with
  unique, non-empty names, and allows at most 100 columns, 50,000 data rows (more sets
  `truncated`) and 8,192 characters a field. A row with a field over the limit, a field count
  other than the header's, or broken quoting is skipped and counted; an unterminated quote fails
  the report (reason `csv`). Values stay strings exactly as written, and callers map them by
  header name.
- Modules and audits reach it through `ScopedGraph.getReport`, with the same declared-pack check
  and stop signal as their other reads.

## Consequences

- One reviewed path follows one redirect. `request()` and `getCollection()` still refuse every
  3xx, including for this pack, and tests hold both.
- No new permission. `Reports.Read.All` already covered these reports, and per-user detail still
  needs Reports Reader or a workload admin role. The `directory.users.read` pack also gains
  `/users/{id}/memberOf` (direct memberships, not transitive) for the audits that follow, since
  `User.Read.All` is Microsoft's documented least privilege for another user's memberships. Its
  id takes only the characters an id or an encoded principal name holds, and no encoded slash,
  backslash or question mark.
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
packs, and `request()` refusing a 302 for every pack. `packages/graph/test/csv.test.ts` covers the
parser's rules and limits and a 10 MB report in well under a second. The synthetic tenant serves
one report through the same redirect (`scripts/demo/graph/usage-report-download.test.ts`) and
refuses a download that carries a token.

## Revisit when

- Microsoft documents a different download host, a JSON form in v1.0 for these reports, or a
  stated encoding.
- A report outgrows 32 MiB or 50,000 rows in a supported tenant.
- Any other feature asks to follow a redirect: it gets its own decision, not a wider list here.
