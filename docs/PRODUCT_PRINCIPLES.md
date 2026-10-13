# Product principles

## Explain before asking

Show what a feature does, which data it needs, which permission enables it, and what happens
when permission is denied—before opening a consent screen.

## Facts before conclusions

Keep observed evidence distinct from interpretation. A finding states its source, time
window, freshness, rule, and limitations. It never presents an audit trail as ground truth
about human intent or performance.

## Safe by default

Start disconnected, local, synthetic, read-only, and minimally retained. Broader access is
an explicit feature decision. Destructive convenience is not a feature.

## Fast paths with guardrails

Administrators need speed during incidents. Good defaults, search, saved views, keyboard
navigation, previews, and clear diffs reduce effort without removing accountability.

## Progressive capability

Prism365 should remain useful when a tenant declines an optional permission. Enable
features independently and report missing capability precisely.

## Transparent operations

Show freshness, sync state, throttling, coverage, provider request IDs, and data gaps in
language an administrator can act on. Do not hide uncertainty behind a green status.

## Governed intelligence

AI is an optional analysis interface, not a privileged identity. Administrators choose the
provider and exact datasets it may receive. Model input is tenant-bound, minimized, bounded,
and auditable; generated conclusions cite supplied evidence and remain untrusted. A model never
silently gains future data or administrative authority.

## Open implementation, original product

Learn from administrator problems and public platform documentation, not proprietary
implementation details. All report logic and interface assets must be original and
reviewable in the repository.

## Accessible is complete

A feature is not done until it works with keyboard navigation, screen readers, visible
focus, reduced motion, sufficient contrast, zoom, and narrow screens.
