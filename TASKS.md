# Agent coordination board

This board coordinates local parallel work before GitHub issues are available. Do not add
secrets, personal data, or customer data. Once published, link each row to a GitHub issue.

| ID | State | Owner | Scope / files | Dependencies | Handoff |
|---|---|---|---|---|---|
| TL-001 | ready | — | Entra app connection spike | ADR-0002, human test tenant | — |
| TL-002 | ready | — | Synthetic user inventory adapter | TL-001 | — |
| TL-003 | ready | — | Sign-in summary adapter | TL-001 | — |
| TL-004 | ready | — | Usage reports adapter | TL-001 | — |
| TL-005 | ready | — | PostgreSQL tenant-isolation design | Threat model review | — |
| TL-006 | ready | — | Dashboard accessibility audit | Running web app | — |
| TL-007 | ready | — | Finding evidence model | Product principles | — |
| TL-008 | ready | — | Self-host deployment profile | TL-005 | — |

## Claim format

Change `ready` to `claimed`, replace the owner, and add a short timestamp/branch note to the
handoff cell. Do not claim more than one substantial task at a time.

## Handoff format

```text
Changed: paths or modules
Verified: commands and manual checks
Permission/data impact: none, or exact change
Risks: known edge cases and assumptions
Next: smallest actionable follow-up
```
