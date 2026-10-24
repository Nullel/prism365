import { randomUUID } from "node:crypto";
import { supportedReadFeatures } from "../../../apps/api/src/config.ts";
import { GRAPH_RESOURCE_APP_ID, SIGN_IN_PERMISSIONS } from "../permissions.ts";

// An invented tenant behind a fake fetch, for the setup's tests. It keeps app registrations,
// service principals and certificates in memory, records every request, and answers the way the
// Microsoft Graph documentation describes (keys only on a single-object $select=keyCredentials,
// PATCH replacing a property as a whole). Nothing here reaches the network.

export const TENANT_ID = "11111111-1111-4111-8111-111111111111";
export const ADMIN_ID = "22222222-2222-4222-8222-222222222222";
export const OTHER_API = "99999999-9999-4999-8999-999999999999";

const names = [
  ...SIGN_IN_PERMISSIONS,
  ...new Set(Object.values(supportedReadFeatures)),
  "Mail.Read",
  "Directory.ReadWrite.All",
];
// Scope IDs as Graph's service principal publishes them; invented, but stable within a run.
export const scopeIds = new Map(
  names.map((name, index) => [
    name,
    `00000000-0000-4000-8000-${String(index + 1).padStart(12, "0")}`,
  ]),
);
export const roleIds = new Map([["Directory.Read.All", "00000000-0000-4000-9000-000000000001"]]);

export interface FakeRequest {
  method: string;
  path: string;
  query: URLSearchParams;
  body: unknown;
}

export interface FakeApplication {
  id: string;
  appId: string;
  displayName: string;
  signInAudience: string;
  tags: string[];
  web: Record<string, unknown> & { redirectUris: string[] };
  requiredResourceAccess: {
    resourceAppId: string;
    resourceAccess: { id: string; type: string }[];
  }[];
  keyCredentials: Record<string, unknown>[];
  passwordCredentials: Record<string, unknown>[];
  publicClient: { redirectUris: string[] };
  spa: { redirectUris: string[] };
  isFallbackPublicClient: boolean | null;
  owners: string[];
  federatedIdentityCredentials: Record<string, unknown>[];
  // Key IDs whose key value the fake leaves out, as Microsoft does on list queries.
  hiddenKeys?: string[];
}

export interface FakeServicePrincipal {
  id: string;
  appId: string;
  tags: string[];
  owners: string[];
  passwordCredentials: Record<string, unknown>[];
  keyCredentials: Record<string, unknown>[];
  appRoleAssignments: { id: string; resourceDisplayName: string }[];
  grants: { id: string; consentType: string; resourceId: string; scope: string }[];
}

// Microsoft Graph's own service principal in the fake tenant.
export const GRAPH_PRINCIPAL_ID = "33333333-3333-4333-8333-333333333333";

const APPLICATION_FIELDS = [
  "id",
  "appId",
  "displayName",
  "signInAudience",
  "tags",
  "web",
  "requiredResourceAccess",
  "passwordCredentials",
  "publicClient",
  "spa",
  "isFallbackPublicClient",
] as const;

export class FakeTenant {
  readonly requests: FakeRequest[] = [];
  applications: FakeApplication[] = [];
  servicePrincipals: FakeServicePrincipal[] = [];
  me = { id: ADMIN_ID, userPrincipalName: "admin@contoso.example" };
  // Answers a request before the fake does, to inject failures.
  intercept: ((request: FakeRequest) => Response | null) | null = null;
  // Key IDs the next keyCredentials PATCH loses, to test the check after it.
  dropOnPatch: string[] = [];
  tokens: string[] = [];

  application(overrides: Partial<FakeApplication> = {}): FakeApplication {
    const application: FakeApplication = {
      id: randomUUID(),
      appId: randomUUID(),
      displayName: "Prism365",
      signInAudience: "AzureADMyOrg",
      tags: [],
      web: {
        redirectUris: [],
        implicitGrantSettings: { enableIdTokenIssuance: false, enableAccessTokenIssuance: false },
      },
      requiredResourceAccess: [],
      keyCredentials: [],
      passwordCredentials: [],
      publicClient: { redirectUris: [] },
      spa: { redirectUris: [] },
      isFallbackPublicClient: null,
      // A registration someone made through the admin centre is owned by them; tests that mean
      // the signed-in administrator's own registration say so.
      owners: [ADMIN_ID],
      federatedIdentityCredentials: [],
      ...overrides,
    };
    this.applications.push(application);
    return application;
  }

  servicePrincipal(appId: string, overrides: Partial<FakeServicePrincipal> = {}) {
    const principal: FakeServicePrincipal = {
      id: randomUUID(),
      appId,
      tags: [],
      owners: [],
      passwordCredentials: [],
      keyCredentials: [],
      appRoleAssignments: [],
      grants: [],
      ...overrides,
    };
    this.servicePrincipals.push(principal);
    return principal;
  }

  get writes() {
    return this.requests.filter((request) => request.method !== "GET");
  }

  readonly fetch = async (input: string | URL, init?: RequestInit): Promise<Response> => {
    const url = new URL(String(input));
    if (url.origin !== "https://graph.microsoft.com" || !url.pathname.startsWith("/v1.0/"))
      throw new Error(`Unexpected host ${url.origin}`);
    this.tokens.push(new Headers(init?.headers).get("Authorization") ?? "");
    const request: FakeRequest = {
      method: init?.method ?? "GET",
      path: decodeURIComponent(url.pathname.slice("/v1.0/".length)),
      query: url.searchParams,
      body: typeof init?.body === "string" ? JSON.parse(init.body) : undefined,
    };
    this.requests.push(request);
    return this.intercept?.(request) ?? this.#answer(request);
  };

  #answer(request: FakeRequest): Response {
    const { method, path, query } = request;
    const filter = query.get("$filter") ?? "";
    if (method === "GET" && path === "me") return Response.json(this.me);
    if (method === "GET" && path === `servicePrincipals(appId='${GRAPH_RESOURCE_APP_ID}')`)
      return Response.json({
        id: GRAPH_PRINCIPAL_ID,
        oauth2PermissionScopes: [...scopeIds].map(([value, id]) => ({
          id,
          value,
          isEnabled: true,
        })),
        appRoles: [...roleIds].map(([value, id]) => ({ id, value, isEnabled: true })),
      });
    if (method === "GET" && path === "servicePrincipals") {
      const appId = /^appId eq '([^']+)'$/.exec(filter)?.[1];
      return Response.json({
        value: this.servicePrincipals
          .filter((principal) => principal.appId === appId)
          .map(({ id }) => ({ id })),
      });
    }
    if (method === "POST" && path === "servicePrincipals") {
      const body = request.body as { appId: string; tags: string[] };
      if (!this.applications.some((application) => application.appId === body.appId))
        return error(400, "Request_BadRequest");
      const principal = this.servicePrincipal(body.appId, { tags: body.tags });
      return Response.json({ id: principal.id, appId: principal.appId }, { status: 201 });
    }
    const owner = (id: string) => ({ id, displayName: `Owner ${id.slice(0, 8)}` });
    const principalPath =
      /^servicePrincipals\/([^/(]+)(?:\/(appRoleAssignments|owners|oauth2PermissionGrants))?$/.exec(
        path,
      );
    if (method === "GET" && principalPath) {
      const principal = this.servicePrincipals.find((item) => item.id === principalPath[1]);
      if (!principal) return error(404, "Request_ResourceNotFound");
      if (principalPath[2] === "appRoleAssignments")
        return Response.json({ value: principal.appRoleAssignments });
      if (principalPath[2] === "owners")
        return Response.json({ value: principal.owners.map(owner) });
      if (principalPath[2] === "oauth2PermissionGrants")
        return Response.json({ value: principal.grants });
      return Response.json({
        passwordCredentials: principal.passwordCredentials,
        keyCredentials: principal.keyCredentials,
      });
    }
    const applicationPath = /^applications\/([^/]+)\/(owners|federatedIdentityCredentials)$/.exec(
      path,
    );
    if (method === "GET" && applicationPath) {
      const application = this.applications.find((item) => item.id === applicationPath[1]);
      if (!application) return error(404, "Request_ResourceNotFound");
      return Response.json({
        value:
          applicationPath[2] === "owners"
            ? application.owners.map(owner)
            : application.federatedIdentityCredentials,
      });
    }
    if (method === "GET" && path === "applications") {
      const tag = /^tags\/any\(t:t eq '([^']+)'\)$/.exec(filter)?.[1];
      const name = /^displayName eq '(.+)'$/.exec(filter)?.[1]?.replaceAll("''", "'");
      const found = this.applications.filter((application) =>
        tag !== undefined ? application.tags.includes(tag) : application.displayName === name,
      );
      return Response.json({ value: found.map(listed) });
    }
    const byAppId = /^applications\(appId='([^']+)'\)$/.exec(path)?.[1];
    if (method === "GET" && byAppId) {
      const application = this.applications.find((item) => item.appId === byAppId);
      return application
        ? Response.json(listed(application))
        : error(404, "Request_ResourceNotFound");
    }
    if (method === "POST" && path === "applications") {
      const body = request.body as Partial<FakeApplication>;
      // Microsoft makes the signed-in caller the owner of a registration it creates.
      const application = this.application({
        ...body,
        id: randomUUID(),
        appId: randomUUID(),
        keyCredentials: [],
        owners: [this.me.id],
      } as Partial<FakeApplication>);
      return Response.json(listed(application), { status: 201 });
    }
    const id = /^applications\/([^/]+)$/.exec(path)?.[1];
    const application = this.applications.find((item) => item.id === id);
    if (id && !application) return error(404, "Request_ResourceNotFound");
    if (method === "GET" && application && query.get("$select") === "id,appId,tags")
      return Response.json(listed(application));
    if (method === "GET" && application && query.get("$select") === "keyCredentials")
      return Response.json({
        keyCredentials: application.keyCredentials.map((credential) =>
          application.hiddenKeys?.includes(String(credential.keyId))
            ? { ...credential, key: null }
            : credential,
        ),
      });
    if (method === "PATCH" && application) {
      const body = request.body as Record<string, unknown>;
      for (const [key, value] of Object.entries(body)) {
        if (key === "keyCredentials") {
          application.keyCredentials = (value as Record<string, unknown>[])
            .map((credential) => ({ keyId: randomUUID(), ...credential }))
            .filter((credential) => !this.dropOnPatch.includes(String(credential.keyId)));
          this.dropOnPatch = [];
        } else (application as unknown as Record<string, unknown>)[key] = value;
      }
      return new Response(null, { status: 204 });
    }
    return error(404, "Request_ResourceNotFound");
  }
}

function listed(application: FakeApplication) {
  return Object.fromEntries(APPLICATION_FIELDS.map((field) => [field, application[field]]));
}

export function error(status: number, code: string): Response {
  return Response.json(
    { error: { code, message: "Provider text that must never be printed." } },
    { status, headers: { "request-id": "synthetic-request" } },
  );
}
