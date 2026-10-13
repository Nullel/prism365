# Tenant and account compatibility

The compatibility goal is precise: any supported Microsoft Entra work or school tenant can
self-host Prism365 and enable the features for which it has the required cloud support,
administrator consent, calling-identity role, product license, and tenant configuration.

It does **not** mean that one permission makes every feature work in every Microsoft cloud.
Personal Microsoft accounts are not a target because the administration APIs generally do
not support them.

The durable architecture decision is [ADR-0003](decisions/0003-provider-capability-model.md).

## Connection models

### Self-hosted, bring-your-own app registration — first release

Each organization creates a single-tenant Entra application in its own directory, configures
the exact reply URL, selects only the feature packs it needs, and supplies a certificate or
federated workload identity for background collection. This is the simplest trust model and
does not require the project maintainers to operate a shared credential.

Interactive delegated access is useful for sign-in and investigations. Scheduled reporting
uses application access only where necessary. Production client secrets are not the default.

### Hosted multi-tenant service — future

A hosted service would require a verified publisher, operational security program, tenant
isolation, deletion and export controls, incident response, and an admin-consent flow. One
multi-tenant registration cannot span Microsoft cloud boundaries, so each supported national
cloud needs a separate app registration and deployment.

## Effective-access rule

For a capability to be available, all of the following must be true:

```text
implemented
AND supported in selected cloud
AND required tenant service is enabled
AND required product license is present
AND requested permission has consent
AND signed-in user or service principal has the required provider role/scope
AND provider is reachable
```

For delegated calls, an OAuth permission never elevates the signed-in user beyond that user's
own directory role and privileges. For application calls, admin consent is necessary but an
API such as Exchange can additionally require provider-specific RBAC.

## Capability states

The API and UI use explicit states rather than a generic “access denied”:

| State | Meaning | UI behavior |
|---|---|---|
| `available` | All known prerequisites pass | Feature can run |
| `missingConsent` | Required API permission is not granted | Show exact permission and consent guidance |
| `missingRole` | Caller lacks a required Entra or workload role | Show acceptable roles without asking for a broader API scope |
| `missingLicense` | Tenant lacks the required product entitlement | Name the dependency; do not upsell or infer data |
| `unsupportedCloud` | Microsoft does not expose the endpoint in this cloud | Disable only this feature |
| `notEnabled` | Required tenant service, audit, or subscription is off | Show the tenant-side prerequisite |
| `notImplemented` | Prism365 has not shipped the adapter | Never imply a tenant problem |
| `providerUnavailable` | Microsoft is throttling, degraded, or unreachable | Preserve stale data with freshness warning where safe |
| `unknown` | Readiness could not be established | Default to unavailable and explain the missing check |

403, 404, and empty data are not enough to identify these states reliably. Onboarding must run
documented readiness probes and retain only the minimum diagnostics needed to explain the
result.

## Microsoft cloud profiles

| Profile | Authority | Graph | Management Activity | Notes |
|---|---|---|---|---|
| Worldwide | `login.microsoftonline.com` | `graph.microsoft.com` | `manage.office.com` | Broadest API coverage |
| GCC | Global authority | Global Graph | `manage-gcc.office.com` | Endpoint support still requires feature-level verification |
| GCC High | `login.microsoftonline.us` | `graph.microsoft.us` | `manage.office365.us` | Separate app/deployment and reduced API catalog |
| DoD | `login.microsoftonline.us` | `dod-graph.microsoft.us` | `manage.protection.apps.mil` | Separate app/deployment and reduced API catalog |
| China (21Vianet) | `login.chinacloudapi.cn` | `microsoftgraph.chinacloudapi.cn` | Not enabled in the current catalog | Many security/reporting APIs are unavailable |

Tokens and next links must never cross profiles. Cloud roots are code-owned allowlists, not
administrator-provided URLs.

## Verification matrix

“Supported” can be claimed only after automated contract tests and an opt-in live smoke test
pass against a non-production tenant for that cloud. The matrix must include:

1. delegated sign-in and tenant-bound token validation;
2. application authentication with certificate or workload identity;
3. incremental consent and consent revocation;
4. expected missing-role, missing-license, and disabled-service states;
5. pagination, throttling, provider outage, and partial-result behavior;
6. tenant isolation and cross-cloud next-link rejection;
7. deletion, credential rotation, and clean disconnect;
8. report privacy/concealment behavior and no productivity claims.

Until this matrix has real-tenant evidence, the repository must say “targeted” or “planned,”
not “works in every tenant.” The current code contains cloud profiles and safety boundaries,
but production tenant authentication is not implemented yet.

## Primary references

- [Single- and multi-tenant apps](https://learn.microsoft.com/en-us/entra/identity-platform/single-and-multi-tenant-apps)
- [Microsoft identity platform application model](https://learn.microsoft.com/en-us/entra/identity-platform/application-model)
- [Microsoft Graph permission types and effective access](https://learn.microsoft.com/en-us/graph/permissions-overview)
- [Microsoft Graph permission best practices](https://learn.microsoft.com/en-us/graph/best-practices-graph-permission)
- [National-cloud Graph deployments](https://learn.microsoft.com/en-us/graph/deployments)
- [Publisher verification](https://learn.microsoft.com/en-us/entra/identity-platform/publisher-verification-overview)
- [Add application credentials](https://learn.microsoft.com/en-us/entra/identity-platform/how-to-add-credentials?tabs=certificate)
