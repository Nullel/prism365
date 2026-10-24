# Governance

Prism365 is an Apache-2.0 community project. During pre-alpha, `@Nullel` is the founding
maintainer and final reviewer for releases, repository administration, permission changes,
authentication, tenant isolation, retention, security policy, and Microsoft connector scope.
This describes review responsibility, not unilateral ownership of contributor copyright.

## Decision model

- Routine, reversible implementation choices use pull-request consensus.
- Architecture, permissions, data collection, retention, deployment, and governance changes
  require a public proposal and explicit maintainer acceptance.
- Security response may be private until coordinated disclosure is safe.
- When consensus cannot be reached, the founding maintainer makes the documented decision;
  contributors may propose that it be revisited with new evidence.

## Maintainer growth

Maintainers are added based on sustained, trustworthy contributions, respectful review,
security and privacy judgment, and reliable stewardship—not commit count. At least two
maintainers should approve future changes to authentication, write permissions, tenant
isolation, or the release pipeline once the maintainer group reaches three people.

## Releases

Releases are cut from a protected default branch after required checks pass. A release must
include notes, dependency provenance, an SBOM, upgrade/rollback instructions where relevant,
and any permission or migration change. Pre-alpha releases must state that production tenant
use is unsupported.

## Project changes

This governance model should be revisited before the first stable release, when adding a
second maintainer, or if a hosted service is created.
