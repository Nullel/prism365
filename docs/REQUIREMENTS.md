# What is needed to operate Prism365

This is a requirements checklist for a real deployment, not a promise that the pre-alpha
scaffold already connects to a tenant.

## For a self-hosted administrator

- A Microsoft Entra work or school tenant; personal Microsoft accounts are not supported.
- An administrator who can create a single-tenant app registration and grant only the
  selected feature packs. Consent authority and workload roles are separate checks.
- A fixed HTTPS origin and exact redirect URI for production sign-in.
- A certificate or federated workload identity for scheduled collection. Store private key
  material in a managed secret/key service; never in source control or application logs.
- Product licenses and tenant services required by the selected reports, such as Entra ID P2
  for risky-user data or Intune for managed-device inventory.
- PostgreSQL with enforced tenant isolation before durable tenant data is enabled.
- A job runner/queue for throttled, resumable collection; it must not share browser-session
  credentials.
- Encrypted backup, restore, retention, deletion, update, and incident-response procedures.
- A public webhook/lifecycle-notification endpoint only if those optional capabilities are
  enabled; validate origin, client state, replay, expiration, and reauthorization.
- For AI analysis, an operator-configured local, tenant-hosted, or external provider; server-side
  credential storage; egress allowlisting; explicit tenant data grants; and reviewed provider
  retention, training, residency, and incident-response terms.

## For contributors and maintainers

- Synthetic fixtures for every normal, missing-permission, partial, and failure path.
- Opt-in non-production test tenants representing supported clouds and license combinations.
- A current endpoint-to-permission-to-role-to-license-to-cloud manifest for every connector.
- Security review for any new permission, external destination, sensitive field, beta API,
  or management action.
- Dependency scanning, secret scanning, CodeQL, protected reviews, signed release artifacts,
  an SBOM, and a documented upgrade/rollback path before a stable release.
- Microsoft documentation re-verification at release time; cloud support and permissions can
  change independently of Prism365.
- AI provider conformance, citation, privacy, prompt-injection, cost, and model-version regression
  tests before any adapter is marked supported.

## MVP infrastructure

```text
Browser -> application API -> authorization/policy -> job queue -> source adapters
                            -> tenant-isolated PostgreSQL
                            -> immutable security/action ledger
```

The smallest credible deployment is the web/API service, one worker, PostgreSQL, a secret or
key service, HTTPS ingress, and observable backups. Object storage is optional for generated
exports and must inherit tenant isolation and retention controls.

## Readiness endpoint requirements

The onboarding API should return, per feature:

- implementation state and connector version;
- selected Microsoft cloud and endpoint availability;
- required versus granted permissions;
- required role categories without over-collecting directory-role membership;
- license/service prerequisites and whether they can be safely verified;
- last probe time, redacted provider request ID, and a human-readable remediation;
- what data will be collected, its retention, and behavior when the feature is disabled.

Never return access tokens, credentials, raw tenant error bodies, or sensitive object IDs in
diagnostic output.
