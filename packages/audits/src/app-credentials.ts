import { type AuditRow, cleanText, encodeOwnerEntry, type OwnerEntry } from "@prism365/contracts";
import { GraphRequestError, GraphRetryDeferredError } from "@prism365/graph";
import { z } from "zod";
import { MAX_REQUESTS, MAX_SERVICE_PRINCIPALS, RequestBudget } from "./app-consents.ts";
import type { Audit } from "./definition.ts";
import { type PagedRead, readPages } from "./graph-pages.ts";

// Every application registration, with the state of its secrets and certificates, who owns it, who
// may sign in to it and where it sends people back to. Microsoft never returns a secret's value,
// only when it starts and ends. Microsoft Graph v1.0 references, field names checked 2026-11-09:
// https://learn.microsoft.com/en-us/graph/api/application-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/application-list-owners?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/serviceprincipal-list?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/passwordcredential?view=graph-rest-1.0
// https://learn.microsoft.com/en-us/graph/api/resources/keycredential?view=graph-rest-1.0

const applicationContext = {
  feature: "directory.applications.read",
  permission: "Application.Read.All",
} as const;

export const MAX_APPLICATIONS = 2_000;
const DAY = 86_400_000;
const SOON_DAYS = 30;
const LATER_DAYS = 90;
// A secret that has been in place longer than two years is a finding in the published baselines.
const OLD_SECRET_DAYS = 730;
const MANY_SECRETS = 5;
const OWNER_NAMES_SHOWN = 5;
const REDIRECTS_SHOWN = 10;
const MAX_CREDENTIALS = 1_000;
const MAX_REDIRECTS_READ = 300;

export const AUDIENCES = [
  "This tenant",
  "Any organisation",
  "Any organisation or personal account",
  "Personal accounts",
] as const;
export const CREDENTIAL_STATUSES = [
  "Expired",
  "Expires within 30 days",
  "Expires within 90 days",
  "Expired leftovers",
  "No credentials",
  "OK",
] as const;
type Status = (typeof CREDENTIAL_STATUSES)[number];

const AUDIENCE_BY_VALUE: Record<string, (typeof AUDIENCES)[number]> = {
  azureadmyorg: "This tenant",
  azureadmultipleorgs: "Any organisation",
  azureadandpersonalmicrosoftaccount: "Any organisation or personal account",
  personalmicrosoftaccount: "Personal accounts",
};

// Names and ids are provider text: cut to length before cleaning, never after.
const text = z
  .string()
  .transform((value) => cleanText(value.slice(0, 500)))
  .nullish()
  .catch(null)
  .transform((value) => value ?? null);
// The object IDs go into request paths, and the owner path only takes a GUID.
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const guid = z.string().regex(GUID);
const credentialList = z
  .array(z.unknown())
  .max(MAX_CREDENTIALS)
  .nullish()
  .catch(null)
  .transform((value) => value ?? []);
const redirectList = z
  .object({
    redirectUris: z
      .array(z.unknown())
      .nullish()
      .catch(null)
      .transform((value) =>
        (value ?? [])
          .slice(0, MAX_REDIRECTS_READ)
          .flatMap((item) => (typeof item === "string" ? [cleanText(item.slice(0, 500))] : [])),
      ),
  })
  .nullish()
  .catch(null)
  .transform((value) => value?.redirectUris ?? []);

const applicationSchema = z.object({
  id: guid,
  appId: text,
  displayName: text,
  createdDateTime: text,
  signInAudience: text,
  passwordCredentials: credentialList,
  keyCredentials: credentialList,
  web: redirectList,
  spa: redirectList,
  publicClient: redirectList,
});
type Application = z.output<typeof applicationSchema>;

const principalSchema = z.object({
  id: guid,
  appId: text,
  servicePrincipalType: text,
  passwordCredentials: credentialList,
  keyCredentials: credentialList,
});

const credentialSchema = z.object({
  startDateTime: z.string().nullish(),
  endDateTime: z.string().nullish(),
  usage: z.string().nullish(),
  // A public thumbprint, never secret material. It is read only to pair a SAML signing
  // certificate's two key credentials and its password; it is not kept, logged or shown.
  customKeyIdentifier: z.string().max(200).nullish(),
});

const ownerSchema = z.object({
  "@odata.type": text,
  id: z.string().min(1).max(200),
  displayName: text,
  userPrincipalName: text,
});

// A credential's dates in milliseconds. A missing date is null; a date that is not one makes the
// whole credential unreadable, because treating it as "never ends" would be a guess.
export interface Credential {
  start: number | null;
  end: number | null;
}

interface Parsed {
  credential: Credential;
  usage: string;
  thumbprint: string | null;
}

function credentialOf(raw: unknown): Parsed | null {
  const parsed = credentialSchema.safeParse(raw);
  if (!parsed.success) return null;
  const at = (value: string | null | undefined): number | null | "bad" => {
    if (value === null || value === undefined) return null;
    const time = Date.parse(value);
    return Number.isNaN(time) ? "bad" : time;
  };
  const start = at(parsed.data.startDateTime);
  const end = at(parsed.data.endDateTime);
  if (start === "bad" || end === "bad") return null;
  const thumbprint = parsed.data.customKeyIdentifier?.trim().toLowerCase() || null;
  return {
    credential: { start, end },
    usage: (parsed.data.usage ?? "").toLowerCase(),
    thumbprint,
  };
}

interface Credentials {
  passwords: Credential[];
  keys: Credential[];
  unreadable: number;
}

interface CredentialLists {
  passwords: unknown[];
  keys: unknown[];
}

// The credentials of a registration, with those of its enterprise application added. Microsoft
// stores a SAML signing certificate on the enterprise application as two key credentials, one to
// sign and one to verify, plus a password credential for the private key. They are one certificate,
// so it is counted once, by its Verify key; a Sign key whose thumbprint matches a Verify key is
// skipped, and so is a password whose thumbprint matches any key of the enterprise application.
// A Sign key with no Verify twin, and a password that pairs with nothing, count as they are.
function readCredentials(app: CredentialLists, principal?: CredentialLists): Credentials {
  const result: Credentials = { passwords: [], keys: [], unreadable: 0 };
  const take = (raw: unknown, into: Credential[]) => {
    const parsed = credentialOf(raw);
    if (parsed) into.push(parsed.credential);
    else result.unreadable += 1;
  };
  for (const raw of app.passwords) take(raw, result.passwords);
  for (const raw of app.keys) take(raw, result.keys);
  if (!principal) return result;
  const keys: Parsed[] = [];
  for (const raw of principal.keys) {
    const parsed = credentialOf(raw);
    if (parsed) keys.push(parsed);
    else result.unreadable += 1;
  }
  const verified = new Set(
    keys.flatMap((key) => (key.usage === "verify" && key.thumbprint ? [key.thumbprint] : [])),
  );
  const paired = new Set(keys.flatMap((key) => (key.thumbprint ? [key.thumbprint] : [])));
  for (const key of keys) {
    const twin = key.usage === "sign" && key.thumbprint !== null && verified.has(key.thumbprint);
    if (!twin) result.keys.push(key.credential);
  }
  for (const raw of principal.passwords) {
    const parsed = credentialOf(raw);
    if (!parsed) result.unreadable += 1;
    else if (!(parsed.thumbprint && paired.has(parsed.thumbprint)))
      result.passwords.push(parsed.credential);
  }
  return result;
}

const LOOPBACK_HOSTS: ReadonlySet<string> = new Set(["localhost", "127.0.0.1", "[::1]"]);

function schemeAndHost(uri: string): { scheme: string; host: string } | null {
  const match = /^([a-z][a-z0-9+.-]*):\/\/(\[[^\]]*\]|[^/:?#]*)/i.exec(uri.trim());
  return match
    ? { scheme: (match[1] ?? "").toLowerCase(), host: (match[2] ?? "").toLowerCase() }
    : null;
}

// A redirect URI that sends the sign-in response over plain HTTP to a host that is not this
// computer, or that holds a wildcard (which lets the response go to addresses nobody listed).
export function isInsecureRedirect(uri: string): boolean {
  if (uri.includes("*")) return true;
  const parts = schemeAndHost(uri);
  return parts !== null && parts.scheme === "http" && !LOOPBACK_HOSTS.has(parts.host);
}

// A redirect URI that points at the computer the browser runs on: fine while building an app,
// worth a look on one that is in use.
export function isLocalhostRedirect(uri: string): boolean {
  const parts = schemeAndHost(uri);
  return parts !== null && LOOPBACK_HOSTS.has(parts.host);
}

export interface Analysis {
  secrets: number;
  certificates: number;
  expired: number;
  nextExpiry: number | null;
  expiresIn: number | null;
  oldestAge: number | null;
  neverExpires: boolean;
  oldSecret: boolean;
  manySecrets: boolean;
  status: Status;
  // Expired credentials with no active one left, and with one.
  broken: boolean;
  leftover: boolean;
}

// Microsoft gives a secret created as "never expires" an end date far in the future (commonly the
// year 2299). An end date more than a hundred years away is treated as no end date at all.
const ENDLESS_DAYS = 36_525;

// What the credentials of one application say at `now`. A credential is active until its end date
// has passed; an expired one is counted once, as expired, and is left out of everything else.
export function analyseCredentials(
  passwords: readonly Credential[],
  keys: readonly Credential[],
  now: number,
): Analysis {
  const live = (list: readonly Credential[]) =>
    list.filter((credential) => credential.end === null || credential.end >= now);
  const activePasswords = live(passwords);
  const activeKeys = live(keys);
  const active = [...activePasswords, ...activeKeys];
  const total = passwords.length + keys.length;
  const expired = total - active.length;
  const endless = (credential: Credential) =>
    credential.end === null || credential.end - now > ENDLESS_DAYS * DAY;
  const ends = active.flatMap((credential) => (endless(credential) ? [] : [credential.end ?? 0]));
  const nextExpiry = ends.length > 0 ? Math.min(...ends) : null;
  const starts = active.flatMap((credential) =>
    credential.start === null ? [] : [credential.start],
  );
  const older = (credential: Credential) =>
    credential.start !== null && now - credential.start > OLD_SECRET_DAYS * DAY;
  const expiresIn = nextExpiry === null ? null : Math.max(0, Math.ceil((nextExpiry - now) / DAY));
  let status: Status = "OK";
  // Expired only when nothing is left that works: the integration is broken. An expired credential
  // beside an active one is a leftover from a rotation, and ranks below an expiry that is near.
  if (expired > 0 && active.length === 0) status = "Expired";
  else if (expiresIn !== null && expiresIn <= SOON_DAYS) status = "Expires within 30 days";
  else if (expiresIn !== null && expiresIn <= LATER_DAYS) status = "Expires within 90 days";
  else if (expired > 0) status = "Expired leftovers";
  else if (total === 0) status = "No credentials";
  return {
    secrets: activePasswords.length,
    certificates: activeKeys.length,
    expired,
    nextExpiry,
    expiresIn,
    oldestAge:
      starts.length > 0 ? Math.max(0, Math.floor((now - Math.min(...starts)) / DAY)) : null,
    neverExpires: active.some(endless),
    oldSecret:
      activePasswords.some(older) ||
      active.some((credential) => endless(credential) && older(credential)),
    manySecrets: activePasswords.length >= MANY_SECRETS,
    status,
    broken: expired > 0 && active.length === 0,
    leftover: expired > 0 && active.length > 0,
  };
}

const plural = (count: number, one: string, many: string) =>
  `${count.toLocaleString("en-US")} ${count === 1 ? one : many}`;

// Microsoft says what an owner is in @odata.type. A record without it that has a sign-in name is
// an account, since only accounts have one; anything else is not claimed to be one.
function ownerKindOf(owner: z.output<typeof ownerSchema>): OwnerEntry["kind"] {
  const type = (owner["@odata.type"] ?? "").toLowerCase();
  if (type.endsWith(".user")) return "user";
  if (type.endsWith(".serviceprincipal")) return "servicePrincipal";
  if (type === "" && owner.userPrincipalName) return "user";
  return "other";
}

interface Owners {
  count: number;
  // The accounts among them, by name; the others count but are not named.
  names: string[];
  list: OwnerEntry[];
}

function ownersOf(items: readonly unknown[]): { owners: Owners; unreadable: number } {
  const owners: Owners = { count: 0, names: [], list: [] };
  let unreadable = 0;
  for (const raw of items) {
    const parsed = ownerSchema.safeParse(raw);
    if (!parsed.success) {
      unreadable += 1;
      continue;
    }
    const kind = ownerKindOf(parsed.data);
    const name = (parsed.data.displayName ?? parsed.data.userPrincipalName ?? "").trim() || null;
    owners.count += 1;
    owners.list.push({
      id: parsed.data.id,
      name: kind === "user" ? (name?.slice(0, 200) ?? null) : null,
      kind,
    });
    if (kind === "user" && name) owners.names.push(name.slice(0, 200));
  }
  owners.names = [...new Set(owners.names)].sort((a, b) => a.localeCompare(b));
  owners.names.length = Math.min(owners.names.length, OWNER_NAMES_SHOWN);
  return { owners, unreadable };
}

interface Entry {
  app: Application;
  analysis: Analysis;
  servicePrincipalId: string | null;
  audience: (typeof AUDIENCES)[number] | null;
  insecure: string[];
  localhost: string[];
  multiTenant: boolean;
  // Owners are read only for an application that holds a credential or shows another flag.
  wantsOwners: boolean;
  owners: Owners | null;
}

const STATUS_RANK: Record<Status, number> = {
  Expired: 0,
  "Expires within 30 days": 1,
  "Expires within 90 days": 2,
  "Expired leftovers": 3,
  OK: 4,
  "No credentials": 5,
};

const isoOrNull = (value: number | string | null): string | null => {
  if (value === null) return null;
  const time = typeof value === "number" ? value : Date.parse(value);
  return Number.isNaN(time) ? null : new Date(time).toISOString();
};

export const appCredentialsAudit: Audit = {
  definition: {
    id: "apps.credentials",
    version: 1,
    title: "App credentials",
    description:
      "Every application registration with the state of its secrets and certificates: which have expired or expire soon, which never expire or are more than two years old, who owns the app, who can sign in to it and where it sends people back to.",
    category: "Applications",
    subject: { one: "application", many: "applications" },
    columns: [
      { key: "app", label: "Application", kind: "text", empty: "Unnamed application" },
      {
        key: "status",
        label: "Status",
        kind: "enum",
        facet: true,
        help: "The first that applies. Expired: every secret and certificate has passed its end date, so the app has nothing that works. Expires within 30 or 90 days: the soonest end date among the active ones. Expired leftovers: an expired credential is still there beside an active one. No credentials: the app holds no secret or certificate. OK: none of these.",
      },
      {
        key: "audience",
        label: "Who can sign in",
        kind: "enum",
        facet: true,
        empty: "Unknown",
        help: "The accounts the app accepts: this tenant only, any organisation, any organisation or personal Microsoft account, or personal accounts only.",
      },
      { key: "created", label: "Created", kind: "date", empty: "Not recorded" },
      {
        key: "secrets",
        label: "Secrets",
        kind: "number",
        help: "Active secrets (passwords): not yet past their end date. Includes secrets on the enterprise application with the same application ID.",
      },
      {
        key: "certificates",
        label: "Certificates",
        kind: "number",
        help: "Active certificates and keys: not yet past their end date. Includes those on the enterprise application with the same application ID.",
      },
      {
        key: "nextExpiry",
        label: "Next expiry",
        kind: "date",
        empty: "None",
        help: "The soonest end date among the active credentials. Empty when there are none or none of them ends. An end date more than a hundred years away counts as no end.",
      },
      {
        key: "expiresIn",
        label: "Days until expiry",
        kind: "number",
        empty: "None",
        help: "Days from the day the audit ran to the soonest end date among the active credentials. Empty when none is active or none of them ends, so an app whose every credential has expired shows none and its status says Expired.",
      },
      {
        key: "expired",
        label: "Expired credentials",
        kind: "number",
        help: "Secrets and certificates already past their end date. They are counted here and nowhere else.",
      },
      {
        key: "oldestAge",
        label: "Oldest credential",
        kind: "number",
        empty: "None",
        help: "Days since the oldest active credential started.",
      },
      {
        key: "owners",
        label: "Owners",
        kind: "number",
        empty: "Not read",
        help: "People and apps listed as owners of the registration (the first page only). Read only for apps that hold a credential or show another flag, and only while the run's request limit lasts.",
      },
      {
        key: "ownerNames",
        label: "Owner names",
        kind: "list",
        empty: "Not read",
        help: "The first five accounts by name. An owner that is not an account counts but is not named.",
      },
      {
        key: "ownerIds",
        label: "Owners with kinds",
        kind: "list",
        hidden: true,
        internal: true,
        empty: "Not read",
        help: "Every owner as an ID, a name and a kind, so a page can link the owners that are accounts.",
      },
      {
        key: "servicePrincipalId",
        label: "Enterprise application ID",
        kind: "text",
        hidden: true,
        internal: true,
        help: "The ID of the enterprise application with this application ID, which opens the app's page.",
      },
      { key: "appId", label: "Application ID", kind: "text", hidden: true },
      {
        key: "redirects",
        label: "Redirect URIs to check",
        kind: "list",
        empty: "None",
        help: "Plain HTTP to a host other than this computer, wildcards, and addresses on this computer. At most ten.",
      },
    ],
    views: [
      { id: "all", label: "All", tags: [] },
      { id: "expired", label: "Expired", tags: ["expired"] },
      {
        id: "expired-leftover",
        label: "Expired credentials left in place",
        tags: ["expired-leftover"],
      },
      { id: "expiring-30", label: "Expiring within 30 days", tags: ["expiring-30"] },
      { id: "expiring-90", label: "Expiring within 90 days", tags: ["expiring-90"] },
      { id: "no-owner", label: "No owner", tags: ["no-owner"] },
      { id: "old-secret", label: "Secrets older than 2 years", tags: ["old-secret"] },
      { id: "never-expires", label: "Never expires", tags: ["never-expires"] },
      { id: "insecure-redirect", label: "Insecure redirect URIs", tags: ["insecure-redirect"] },
      { id: "multi-tenant", label: "Open to other organisations", tags: ["multi-tenant"] },
      { id: "no-credentials", label: "No credentials", tags: ["no-credentials"] },
      { id: "many-secrets", label: "Five or more secrets", tags: ["many-secrets"] },
      { id: "localhost-redirect", label: "Localhost redirect URIs", tags: ["localhost-redirect"] },
      { id: "owners-not-read", label: "Owners not read", tags: ["owners-not-read"] },
    ],
    defaultView: "all",
    defaultSort: { column: "nextExpiry", direction: "asc" },
    requirements: [
      "A role that can read applications and their owners, such as Directory Readers, Global Reader or Application Administrator.",
    ],
  },
  packs: ["directory.applications.read"],
  optionalPacks: [],
  scoped: false,
  forbidden:
    "Microsoft refused to read applications. It needs Application.Read.All and a role that can read applications and their owners, such as Directory Readers, Global Reader or Application Administrator.",
  async run(context) {
    const budget = new RequestBudget(MAX_REQUESTS);
    const now = context.now.getTime();
    let complete = true;
    const notes: string[] = [];
    const slowDown = (count: number, what: string) =>
      `Microsoft asked Prism365 to slow down after ${plural(count, what, `${what}s`)}, so the rest were not read. Run the audit again in a few minutes.`;

    const applicationList = await readPages(
      context,
      applicationContext,
      "applications?$select=id,appId,displayName,createdDateTime,signInAudience,passwordCredentials,keyCredentials,web,spa,publicClient&$top=100",
      { maxRecords: MAX_APPLICATIONS, allow: () => budget.take() },
    );
    if (!applicationList.complete) {
      complete = false;
      notes.push(
        applicationList.throttled
          ? slowDown(applicationList.items.length, "application")
          : `Stopped after ${plural(applicationList.items.length, "application", "applications")}; the rest of the directory was not read.`,
      );
    }
    let unreadableRecords = 0;
    const applications: Application[] = [];
    const seen = new Set<string>();
    for (const raw of applicationList.items) {
      const parsed = applicationSchema.safeParse(raw);
      if (!parsed.success) {
        unreadableRecords += 1;
        continue;
      }
      const key = parsed.data.id.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      applications.push(parsed.data);
    }

    // The enterprise application carries credentials of its own (a SAML signing certificate, for
    // one) and the ID that opens the app's page. Only the fields the audit uses are asked for.
    const principalList = await readPages(
      context,
      applicationContext,
      "servicePrincipals?$select=id,appId,displayName,servicePrincipalType,passwordCredentials,keyCredentials&$top=100",
      { maxRecords: MAX_SERVICE_PRINCIPALS, allow: () => budget.take() },
    );
    let principalsRead = 0;
    const principalByAppId = new Map<string, z.output<typeof principalSchema>>();
    for (const raw of principalList.items) {
      const parsed = principalSchema.safeParse(raw);
      if (!parsed.success) {
        unreadableRecords += 1;
        continue;
      }
      // Managed identities and the other types do not belong to a registration.
      if (parsed.data.servicePrincipalType?.toLowerCase() !== "application") continue;
      principalsRead += 1;
      const appId = parsed.data.appId?.toLowerCase();
      if (appId && !principalByAppId.has(appId)) principalByAppId.set(appId, parsed.data);
    }
    if (!principalList.complete) {
      complete = false;
      notes.push(
        principalList.throttled
          ? slowDown(principalList.items.length, "service principal")
          : `Stopped after ${plural(principalList.items.length, "service principal", "service principals")}. Credentials held by the rest, and the link to their enterprise application, are missing.`,
      );
    }

    let unreadableCredentials = 0;
    const entries: Entry[] = applications.map((app) => {
      const principal = app.appId ? principalByAppId.get(app.appId.toLowerCase()) : undefined;
      const credentials = readCredentials(
        { passwords: app.passwordCredentials, keys: app.keyCredentials },
        principal && { passwords: principal.passwordCredentials, keys: principal.keyCredentials },
      );
      unreadableCredentials += credentials.unreadable;
      const analysis = analyseCredentials(credentials.passwords, credentials.keys, now);
      const redirects = [...app.web, ...app.spa, ...app.publicClient];
      const insecure = [...new Set(redirects.filter(isInsecureRedirect))].sort();
      const localhost = [
        ...new Set(redirects.filter((uri) => !isInsecureRedirect(uri) && isLocalhostRedirect(uri))),
      ].sort();
      const audience = AUDIENCE_BY_VALUE[(app.signInAudience ?? "").toLowerCase()] ?? null;
      const multiTenant = audience !== null && audience !== "This tenant";
      const holdsCredential = credentials.passwords.length + credentials.keys.length > 0;
      return {
        app,
        analysis,
        servicePrincipalId: principal?.id ?? null,
        audience,
        insecure,
        localhost,
        multiTenant,
        wantsOwners: holdsCredential || insecure.length > 0 || localhost.length > 0 || multiTenant,
        owners: null,
      };
    });

    // Owners, one application at a time, the ones that matter most first so a tenant too large for
    // the request budget loses the quiet applications before the expired ones.
    const queue = entries
      .filter((entry) => entry.wantsOwners)
      .sort(
        (a, b) =>
          STATUS_RANK[a.analysis.status] - STATUS_RANK[b.analysis.status] ||
          (a.app.displayName ?? "")
            .toLowerCase()
            .localeCompare((b.app.displayName ?? "").toLowerCase()) ||
          a.app.id.localeCompare(b.app.id),
      );
    let ownersStopped: "budget" | "throttled" | "refused" | null = null;
    let unreadableOwners = 0;
    for (const entry of queue) {
      context.signal.throwIfAborted();
      let outOfRequests = false;
      let read: PagedRead;
      try {
        read = await readPages(
          context,
          applicationContext,
          `applications/${entry.app.id}/owners?$select=id,displayName,userPrincipalName`,
          {
            maxRecords: 100,
            maxPages: 1,
            allow: () => {
              outOfRequests = !budget.take();
              return !outOfRequests;
            },
          },
        );
      } catch (error) {
        // The registration was deleted after it was listed.
        if (error instanceof GraphRequestError && error.status === 404) {
          entry.wantsOwners = false;
          continue;
        }
        if (error instanceof GraphRetryDeferredError) ownersStopped = "throttled";
        else if (error instanceof GraphRequestError && error.status === 403)
          ownersStopped = "refused";
        else throw error;
        break;
      }
      if (outOfRequests) {
        ownersStopped = "budget";
        break;
      }
      const { owners, unreadable } = ownersOf(read.items);
      unreadableOwners += unreadable;
      entry.owners = owners;
    }
    const ownersRead = entries.filter((entry) => entry.owners !== null).length;
    const ownersMissing = entries.filter((entry) => entry.wantsOwners && entry.owners === null);
    if (ownersMissing.length > 0) {
      complete = false;
      notes.push(
        ownersStopped === "throttled"
          ? `Microsoft asked Prism365 to slow down after the owners of ${ownersRead.toLocaleString("en-US")} of ${plural(queue.length, "application", "applications")} were read. Run the audit again in a few minutes.`
          : ownersStopped === "refused"
            ? "Microsoft refused to list owners, so the remaining applications show Owners not read. Listing owners needs Application.Read.All and a role that can read applications and their owners."
            : `Read the owners of ${ownersRead.toLocaleString("en-US")} of ${plural(queue.length, "application", "applications")} before reaching the limit of ${MAX_REQUESTS} requests per run. Applications with an expired or expiring credential are read first. ${plural(ownersMissing.length, "application shows", "applications show")} Owners not read.`,
      );
    }
    if (unreadableRecords > 0) {
      complete = false;
      notes.push(
        `${plural(unreadableRecords, "record", "records")} could not be read and ${unreadableRecords === 1 ? "is" : "are"} left out.`,
      );
    }
    if (unreadableCredentials > 0 || unreadableOwners > 0) {
      complete = false;
      notes.push(
        `${plural(unreadableCredentials, "credential", "credentials")} and ${plural(unreadableOwners, "owner record", "owner records")} could not be read and ${unreadableCredentials + unreadableOwners === 1 ? "is" : "are"} left out of the counts.`,
      );
    }

    const rows: AuditRow[] = entries.map((entry) => {
      const { analysis, owners } = entry;
      const tags = [
        ...(analysis.broken ? ["expired"] : []),
        ...(analysis.leftover ? ["expired-leftover"] : []),
        ...(analysis.expiresIn !== null && analysis.expiresIn <= SOON_DAYS ? ["expiring-30"] : []),
        ...(analysis.expiresIn !== null && analysis.expiresIn <= LATER_DAYS ? ["expiring-90"] : []),
        ...(analysis.neverExpires ? ["never-expires"] : []),
        ...(analysis.oldSecret ? ["old-secret"] : []),
        ...(analysis.manySecrets ? ["many-secrets"] : []),
        ...(analysis.status === "No credentials" ? ["no-credentials"] : []),
        ...(owners !== null && owners.count === 0 ? ["no-owner"] : []),
        ...(entry.wantsOwners && owners === null ? ["owners-not-read"] : []),
        ...(entry.insecure.length > 0 ? ["insecure-redirect"] : []),
        ...(entry.localhost.length > 0 ? ["localhost-redirect"] : []),
        ...(entry.multiTenant ? ["multi-tenant"] : []),
      ];
      return {
        key: entry.app.id,
        personId: null,
        tags,
        cells: {
          app: entry.app.displayName,
          audience: entry.audience,
          created: isoOrNull(entry.app.createdDateTime),
          secrets: analysis.secrets,
          certificates: analysis.certificates,
          nextExpiry: isoOrNull(analysis.nextExpiry),
          expiresIn: analysis.expiresIn,
          expired: analysis.expired,
          oldestAge: analysis.oldestAge,
          owners: owners?.count ?? null,
          ownerNames: owners?.names ?? null,
          ownerIds: owners ? owners.list.map(encodeOwnerEntry) : null,
          servicePrincipalId: entry.servicePrincipalId,
          appId: entry.app.appId,
          redirects: [...entry.insecure, ...entry.localhost].slice(0, REDIRECTS_SHOWN),
          status: analysis.status,
        },
      };
    });
    rows.sort((a, b) =>
      String(a.cells.app ?? "")
        .toLowerCase()
        .localeCompare(String(b.cells.app ?? "").toLowerCase()),
    );

    const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
    const tone = (value: number, level: "warning" | "critical") => (value > 0 ? level : "good");
    // Until every owner was read, a zero is not an answer.
    const ownersKnown = ownersMissing.length === 0;

    return {
      rows,
      complete,
      summary: [
        {
          label: "Expired credentials",
          value: count("expired"),
          tone: tone(count("expired"), "critical"),
          view: "expired",
        },
        {
          label: "Expired credentials left in place",
          value: count("expired-leftover"),
          tone: "neutral",
          view: "expired-leftover",
        },
        {
          label: "Expiring within 30 days",
          value: count("expiring-30"),
          tone: tone(count("expiring-30"), "warning"),
          view: "expiring-30",
        },
        {
          label: "No owner",
          value: count("no-owner"),
          tone:
            ownersKnown || count("no-owner") > 0 ? tone(count("no-owner"), "warning") : "neutral",
          view: "no-owner",
        },
        {
          label: "Secrets older than 2 years",
          value: count("old-secret"),
          tone: tone(count("old-secret"), "warning"),
          view: "old-secret",
        },
        {
          label: "Insecure redirect URIs",
          value: count("insecure-redirect"),
          tone: tone(count("insecure-redirect"), "warning"),
          view: "insecure-redirect",
        },
        { label: "Applications", value: rows.length, tone: "neutral", view: "all" },
      ],
      notes: [
        `Read ${plural(applications.length, "application", "applications")} and ${plural(principalsRead, "enterprise application", "enterprise applications")}. Credentials on an enterprise application are added to the registration with the same application ID.`,
        `Owners were read for ${ownersRead.toLocaleString("en-US")} of ${plural(applications.length, "application", "applications")}: those that hold a credential or show another flag, within a limit of ${MAX_REQUESTS} requests per run. Only the first page of owners is read.`,
        "Microsoft never returns the value of a secret or certificate, only when it starts and ends. Prism365 cannot tell which credential an integration actually uses.",
        "Active means not yet past its end date. An expired credential is counted as expired and nowhere else. The tiles count applications, not credentials.",
        `A secret counts as old when it is active and started more than ${OLD_SECRET_DAYS} days ago, or when a credential that never ends started that long ago. Five or more active secrets are flagged as many.`,
        "A redirect URI is flagged as insecure when it uses plain HTTP to a host other than this computer, or holds a wildcard. Addresses on this computer are listed separately as localhost.",
        ...notes,
      ],
    };
  },
};
