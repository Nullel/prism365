# ADR-0002: Microsoft identity and tenant connection model

- Status: proposed
- Date: 2026-10-13
- Owners: identity/security maintainer
- Related issue: TL-001 / Milestone 1

## Context

Self-hosted installations may prefer an app registration they control, while a future hosted
service may use a verified multi-tenant app. Interactive administration and background
collection have different authorization risks. National clouds use different issuers and
endpoints. Choosing prematurely could bake broad application permissions into every install.

## Candidate decision

For the first test-tenant milestone:

- support one customer-controlled single-tenant app registration;
- use delegated `User.Read` for administrator sign-in;
- use a separate confidential client only for explicitly enabled background read packs;
- prefer a certificate or workload identity over a long-lived client secret;
- store server-side sessions and tokens only; never place provider tokens in browser storage;
- validate issuer, audience, tenant, nonce/state, and authorized local role;
- keep tenant writes disabled and preserve a distinct future write identity;
- support the global Microsoft cloud first and fail closed for other clouds.

This is a proposal, not an implemented or accepted connection contract.

## Consequences

- Self-hosters retain control of tenant consent and credentials.
- Initial onboarding requires more setup than a shared hosted app.
- Background read access remains visible and separable from interactive login.
- Sovereign-cloud support is deferred until its issuer and endpoint behavior is tested.

## Validation

Use an opt-in non-production tenant. Test consent denial/revocation, role removal, token
expiry, issuer mismatch, cross-tenant requests, secret redaction, and clean disconnect.

## Revisit when

A hosted profile, managed identity deployment, sovereign cloud, or verified multi-tenant app
is proposed.
