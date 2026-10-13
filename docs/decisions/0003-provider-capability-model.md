# ADR-0003: model providers and prerequisites per capability

- Status: accepted
- Date: 2026-10-13
- Owners: maintainers

## Context

Microsoft 365 administration spans Microsoft Graph, Purview Audit, the Office 365 Management
Activity API, and workload-specific surfaces. OAuth consent alone does not determine access:
the selected Microsoft cloud, calling identity's role, product license, tenant settings, and
endpoint availability can all independently disable a feature.

Treating every failure as “missing permission” would encourage broader consent, mislead
administrators, and make national-cloud support unsafe.

## Decision

- Implement each product capability as an opt-in permission/source pack.
- Keep identity and provider endpoints in code-owned cloud profiles.
- Represent missing consent, role, license, service enablement, cloud support, implementation,
  and provider availability as distinct states.
- Start self-hosting with a customer-owned single-tenant app registration.
- Require a separate deployment and app registration for every national-cloud boundary.
- Keep Graph, Purview, Management Activity, and workload-specific connectors isolated even
  where their records overlap.
- Make unsupported features degrade independently; never replace them with a broader source
  or permission without explicit consent and review.

## Consequences

Onboarding and testing are more involved, and complete coverage needs multiple test tenants.
In return, administrators can understand exactly what a feature needs, safely run a subset of
the product, revoke it independently, and avoid sending national-cloud tokens or continuation
links to the wrong host.
