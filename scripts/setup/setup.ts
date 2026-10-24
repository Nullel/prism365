import { randomBytes, X509Certificate } from "node:crypto";
import { readFile } from "node:fs/promises";
import { join, relative } from "node:path";
import type { z } from "zod";
import { isUuid } from "../../apps/api/src/config.ts";
import { type AccessList, type AclTools, setfaclCommand, systemAcl } from "./acl.ts";
import {
  certificateFiles,
  certificateKey,
  certificateRequest,
  inspectSecrets,
  newCredential,
  promotePending,
  readRecord,
  selfSignedCertificate,
  setMode,
  thumbprintOf,
  writeCredential,
  writePending,
  writeRecord,
} from "./certificate.ts";
import { mergeEnv, readEnv, writeEnvFile } from "./env-file.ts";
import {
  applicationListSchema,
  applicationSchema,
  appRoleAssignmentListSchema,
  explainGraphError,
  type Fetch,
  type GraphApplication,
  graphServicePrincipalSchema,
  keyCredentialsResponseSchema,
  looseListSchema,
  meSchema,
  ownerListSchema,
  permissionGrantListSchema,
  SetupGraph,
  SetupGraphError,
  servicePrincipalCredentialsSchema,
  servicePrincipalListSchema,
  servicePrincipalSchema,
} from "./graph.ts";
import {
  hasCertificate,
  mergeKeyCredentials,
  type NewKeyCredential,
  verifyKeyCredentials,
} from "./key-credentials.ts";
import type { SetupOptions } from "./options.ts";
import { GRAPH_RESOURCE_APP_ID } from "./permissions.ts";
import {
  buildPlan,
  consentLink,
  managedValues,
  planText,
  type RegistrationFacts,
  RUN_TAG_PREFIX,
  type SetupPlan,
  type SetupState,
} from "./plan.ts";

export interface SignedIn {
  // The tenant's ID as Microsoft reported it for the signed-in account.
  tenantId: string;
  getToken: () => Promise<string>;
}

export interface SetupDependencies {
  fetch: Fetch;
  signIn: (tenant: string) => Promise<SignedIn>;
  confirm: (question: string) => Promise<boolean>;
  print: (text: string) => void;
  envExample: string;
  cwd: string;
  now?: () => Date;
  sleep?: (milliseconds: number) => Promise<void>;
  acl?: AclTools;
  platform?: NodeJS.Platform;
}

export type SetupOutcome = "applied" | "unchanged" | "dry-run" | "declined";

// A failure the administrator can act on, already in plain words.
export class SetupError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SetupError";
  }
}

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
].join(",");
// A registration made moments ago can take a little while to be visible everywhere in Microsoft
// Entra. Steps on it wait and retry, about 45 seconds in all.
const NEW_OBJECT_DELAYS = [2_000, 4_000, 8_000, 15_000, 15_000];
// After a service principal create that may or may not have happened, wait this long and look
// for it before trying again.
const UNCERTAIN_CREATE_DELAYS = [5_000, 15_000];

export const CLI_TOOLS_REMINDER =
  "Microsoft keeps the access you gave the Microsoft Graph command line tools to sign this setup in. Revoke it now that setup is done: docs/SETUP.md, Revoke the setup's access.";

async function step<T>(label: string, work: () => Promise<T>): Promise<T> {
  try {
    return await work();
  } catch (error) {
    if (error instanceof SetupGraphError) throw new SetupError(explainGraphError(label, error));
    throw error;
  }
}

async function readText(path: string): Promise<string | null> {
  try {
    return await readFile(path, "utf8");
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw error;
  }
}

async function readEnvSource(options: SetupOptions, example: string) {
  const existing = await readText(options.envFile);
  if (existing !== null) return { exists: true, text: existing, source: "env" as const };
  const template = await readText(example);
  return template !== null
    ? { exists: false, text: template, source: "example" as const }
    : { exists: false, text: "", source: "empty" as const };
}

// Every page of a directory-object list, up to a bound.
async function allPages<T>(
  graph: SetupGraph,
  first: string,
  schema: z.ZodType<{ value: T[]; "@odata.nextLink"?: string | undefined }>,
): Promise<{ items: T[]; complete: boolean }> {
  const items: T[] = [];
  let next: string | undefined = first;
  for (let page = 0; next && page < 20; page += 1) {
    const result: { value: T[]; "@odata.nextLink"?: string | undefined } = await graph.get(
      next,
      schema,
    );
    items.push(...result.value);
    next = result["@odata.nextLink"];
  }
  return { items, complete: !next };
}

async function readOwners(graph: SetupGraph, path: string) {
  const { items, complete } = await allPages(graph, path, ownerListSchema);
  return {
    owners: items.map((owner) => ({
      id: owner.id,
      name: owner.userPrincipalName ?? owner.displayName ?? null,
    })),
    complete,
  };
}

async function servicePrincipalOf(graph: SetupGraph, appId: string) {
  const list = await graph.get(
    `servicePrincipals?$filter=${encodeURIComponent(`appId eq '${appId}'`)}&$select=id`,
    servicePrincipalListSchema,
  );
  return list.value[0]?.id ?? null;
}

async function readAcl(acl: AclTools, path: string): Promise<AccessList | null> {
  try {
    return await acl.read(path);
  } catch {
    return null;
  }
}

// The registration to work on. The setup never searches for one: a tag or a name can be put on
// any registration by anyone allowed to register apps. It is the one named by --app-id, or the
// one this install already uses (ENTRA_CLIENT_ID in the .env it writes), or none: a new one.
async function findTarget(graph: SetupGraph, options: SetupOptions, envText: string) {
  const fromEnv = readEnv(envText).get("ENTRA_CLIENT_ID")?.trim().toLowerCase() ?? "";
  const none = { application: null, foundBy: null, missingTarget: null, invalidEnvClientId: null };
  if (!options.appId && fromEnv && !isUuid(fromEnv))
    return { ...none, invalidEnvClientId: fromEnv };
  const appId = options.appId ?? (fromEnv || null);
  if (!appId) return none;
  const source = options.appId ? ("appId" as const) : ("env" as const);
  try {
    const application = await graph.get(
      `applications(appId='${appId}')?$select=${APPLICATION_FIELDS}`,
      applicationSchema,
    );
    return { ...none, application, foundBy: source };
  } catch (error) {
    if (error instanceof SetupGraphError && error.status === 404)
      return { ...none, missingTarget: { appId, source } };
    throw error;
  }
}

async function readRegistration(
  graph: SetupGraph,
  application: GraphApplication,
): Promise<RegistrationFacts> {
  const owners = await readOwners(graph, `applications/${application.id}/owners`);
  const federated = await allPages(
    graph,
    `applications/${application.id}/federatedIdentityCredentials`,
    looseListSchema,
  );
  const principal = await servicePrincipalOf(graph, application.appId);
  if (!principal)
    return {
      owners: owners.owners,
      ownersComplete: owners.complete,
      federatedCredentials: federated.items.length,
      servicePrincipal: null,
    };
  const principalOwners = await readOwners(graph, `servicePrincipals/${principal}/owners`);
  const credentials = await graph.get(
    `servicePrincipals/${principal}?$select=passwordCredentials,keyCredentials`,
    servicePrincipalCredentialsSchema,
  );
  const assignments = await allPages(
    graph,
    `servicePrincipals/${principal}/appRoleAssignments?$select=id,resourceDisplayName`,
    appRoleAssignmentListSchema,
  );
  const grants = await allPages(
    graph,
    `servicePrincipals/${principal}/oauth2PermissionGrants`,
    permissionGrantListSchema,
  );
  return {
    owners: owners.owners,
    ownersComplete: owners.complete,
    federatedCredentials: federated.items.length,
    servicePrincipal: {
      id: principal,
      owners: principalOwners.owners,
      ownersComplete: principalOwners.complete,
      passwordCredentials: credentials.passwordCredentials?.length ?? 0,
      keyCredentials: credentials.keyCredentials ?? [],
      appRoleAssignments: assignments.items,
      grants: grants.items,
    },
  };
}

export async function gatherState(
  options: SetupOptions,
  graph: SetupGraph,
  tenantId: string,
  dependencies: Pick<SetupDependencies, "envExample" | "cwd"> & {
    now: Date;
    acl: AclTools;
    platform: NodeJS.Platform;
  },
): Promise<SetupState> {
  const admin = await step("Reading the signed-in account", () =>
    graph.get("me?$select=id,userPrincipalName", meSchema),
  );
  const env = await readEnvSource(options, dependencies.envExample);
  const target = await step("Reading the registration to use", () =>
    findTarget(graph, options, env.source === "env" ? env.text : ""),
  );
  const graphPrincipal = await step("Reading Microsoft Graph's permission list", () =>
    graph.get(
      `servicePrincipals(appId='${GRAPH_RESOURCE_APP_ID}')?$select=id,oauth2PermissionScopes,appRoles`,
      graphServicePrincipalSchema,
    ),
  );
  let keyCredentials: unknown[] = [];
  let registration: RegistrationFacts | null = null;
  const application = target.application;
  if (application) {
    registration = await step(
      "Checking the registration's owners, credentials and permissions",
      () => readRegistration(graph, application),
    );
    keyCredentials = await step("Reading the registration's certificates", async () => {
      const response = await graph.get(
        `applications/${application.id}?$select=keyCredentials`,
        keyCredentialsResponseSchema,
      );
      return response.keyCredentials;
    });
  }
  const files = certificateFiles(options.secretsDir);
  const secrets = await inspectSecrets(files);
  const credential = secrets.credential;
  return {
    options,
    tenantId,
    admin: { id: admin.id, userPrincipalName: admin.userPrincipalName ?? null },
    application,
    foundBy: target.foundBy,
    missingTarget: target.missingTarget,
    invalidEnvClientId: target.invalidEnvClientId,
    registration,
    servicePrincipal: registration?.servicePrincipal != null,
    graph: graphPrincipal,
    keyCredentials,
    files,
    secrets,
    record: await readRecord(files),
    acl: {
      directory:
        secrets.directoryMode === null ? null : await readAcl(dependencies.acl, files.directory),
      privateKey:
        credential.state === "pair" || credential.state === "keyOnly"
          ? await readAcl(dependencies.acl, files.privateKey)
          : null,
      certificate:
        credential.state === "pair" ? await readAcl(dependencies.acl, files.certificate) : null,
    },
    env,
    cwd: dependencies.cwd,
    now: dependencies.now,
    platform: dependencies.platform,
  };
}

function changesNothing(plan: SetupPlan) {
  return (
    plan.application.action === "keep" &&
    plan.servicePrincipal === "keep" &&
    plan.certificate.action === "keep" &&
    plan.restrictDirectory === null &&
    plan.restrictPrivateKey === null &&
    plan.reapplyAcl === null &&
    !plan.env.write
  );
}

async function retryWhileNew<T>(
  fresh: boolean,
  sleep: (milliseconds: number) => Promise<void>,
  work: () => Promise<T>,
  retryable: (error: SetupGraphError) => boolean,
): Promise<T> {
  const delays = fresh ? NEW_OBJECT_DELAYS : [];
  for (let attempt = 0; ; attempt += 1) {
    try {
      return await work();
    } catch (error) {
      const delay = delays[attempt];
      if (delay === undefined || !(error instanceof SetupGraphError) || !retryable(error))
        throw error;
      await sleep(delay);
    }
  }
}

// A create answered with a server error, a throttle or nothing at all may still have happened.
const uncertain = (error: unknown) =>
  !(error instanceof SetupGraphError) || error.status >= 500 || error.status === 429;

// The registrations labelled with this run's marker that the administrator owns. Only this run
// can know the marker, and a copy someone else made would not be theirs.
async function findByRunLabel(graph: SetupGraph, marker: string, adminId: string) {
  const page = await graph.get(
    `applications?$filter=${encodeURIComponent(`tags/any(t:t eq '${marker}')`)}&$select=id,appId`,
    applicationListSchema,
  );
  const mine: Pick<GraphApplication, "id" | "appId">[] = [];
  for (const application of page.value) {
    const { owners } = await readOwners(graph, `applications/${application.id}/owners`);
    if (owners.some((owner) => owner.id.toLowerCase() === adminId.toLowerCase()))
      mine.push({ id: application.id, appId: application.appId });
  }
  return mine;
}

// Creates the registration, labelled with a marker unique to this run. When Microsoft's answer
// leaves it unclear whether the create happened (a server error, a throttle, no answer), the setup
// looks for that marker for as long as Entra may take to show a new object, and sends the create
// again only once the registration is confirmed absent after that whole wait. A registration that
// takes even longer to appear could still end up twice, so the failure message names the marker.
async function createApplication(
  graph: SetupGraph,
  body: Record<string, unknown>,
  adminId: string,
  sleep: (milliseconds: number) => Promise<void>,
): Promise<{ application: Pick<GraphApplication, "id" | "appId">; marker: string }> {
  const marker = `${RUN_TAG_PREFIX}${randomBytes(8).toString("hex")}`;
  const labelled = { ...body, tags: [...((body.tags as string[] | undefined) ?? []), marker] };
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      return { application: await graph.post("applications", labelled, applicationSchema), marker };
    } catch (error) {
      if (!uncertain(error)) throw error;
      for (const delay of NEW_OBJECT_DELAYS) {
        await sleep(delay);
        const found = await findByRunLabel(graph, marker, adminId);
        if (found.length > 1)
          throw new SetupError(
            `Microsoft made ${found.length} registrations labelled ${marker}. Delete all but one in the Entra admin center, then run the setup again with --app-id for the one you keep.`,
          );
        const [made] = found;
        if (made) return { application: made, marker };
      }
    }
  }
  throw new SetupError(
    `Microsoft did not confirm the new registration. It may still appear: in the Entra admin center, look for a registration with the label ${marker}. If you find it, run the setup again with --app-id and its application ID; otherwise run it again as before.`,
  );
}

// Takes this run's marker (and any an interrupted run left) off the registration. Not fatal.
async function removeRunLabels(graph: SetupGraph, applicationId: string) {
  const current = await graph.get(
    `applications/${applicationId}?$select=id,appId,tags`,
    applicationSchema,
  );
  const tags = current.tags ?? [];
  const kept = tags.filter((tag) => !tag.startsWith(RUN_TAG_PREFIX));
  if (kept.length !== tags.length)
    await graph.patch(`applications/${applicationId}`, { tags: kept });
}

async function createServicePrincipal(
  graph: SetupGraph,
  appId: string,
  fresh: boolean,
  sleep: (milliseconds: number) => Promise<void>,
) {
  let uncertainTries = 0;
  for (let attempt = 0; ; attempt += 1) {
    try {
      await graph.post(
        "servicePrincipals",
        { appId, tags: ["WindowsAzureActiveDirectoryIntegratedApp"] },
        servicePrincipalSchema,
      );
      return;
    } catch (error) {
      if (uncertain(error)) {
        const delay = UNCERTAIN_CREATE_DELAYS[uncertainTries];
        uncertainTries += 1;
        if (delay === undefined) throw error;
        await sleep(delay);
        if (await servicePrincipalOf(graph, appId)) return;
        continue;
      }
      // A 400 or 404 means it was not created; on a registration made moments ago that is usually
      // Entra not having caught up yet.
      const delay = fresh ? NEW_OBJECT_DELAYS[attempt] : undefined;
      const definite = error as SetupGraphError;
      if (delay === undefined || (definite.status !== 400 && definite.status !== 404)) throw error;
      await sleep(delay);
    }
  }
}

const keyIds = (list: readonly unknown[]) =>
  list
    .map((item) => String((item as { keyId?: unknown }).keyId ?? "").toLowerCase())
    .sort()
    .join(",");

// Adds the certificate to the registration without touching any other. A PATCH replaces the whole
// list, so the list is read right before it and compared with the one the plan was made from: if
// someone changed the certificates in between, the setup stops rather than undo their change.
// After the PATCH the list is read again, and every certificate from right before must be there.
// (Microsoft's addKey would add without replacing, but it needs proof signed by a certificate
// already on the registration, which the first certificate cannot have, and its proof format can
// only be checked against a real tenant; the PATCH path has to be safe either way.)
async function addCertificate(
  graph: SetupGraph,
  applicationId: string,
  certificate: X509Certificate,
  planned: readonly unknown[],
  fresh: boolean,
  sleep: (milliseconds: number) => Promise<void>,
): Promise<"added" | "present"> {
  const key = certificateKey(certificate);
  const read = () =>
    graph
      .get(`applications/${applicationId}?$select=keyCredentials`, keyCredentialsResponseSchema)
      .then((response) => response.keyCredentials);
  const before = await step("Reading the registration's certificates", () =>
    retryWhileNew(fresh, sleep, read, (error) => error.status === 404),
  );
  if (hasCertificate(before, key)) return "present";
  if (keyIds(before) !== keyIds(planned))
    throw new SetupError(
      "The registration's certificates changed while the setup was running (someone added or removed one). Its certificates were not changed. Run the setup again once nobody is editing them.",
    );
  const added: NewKeyCredential = {
    type: "AsymmetricX509Cert",
    usage: "Verify",
    key,
    displayName: certificate.subject.replace(/\n/g, ", "),
    startDateTime: certificate.validFromDate.toISOString(),
    endDateTime: certificate.validToDate.toISOString(),
  };
  const list = mergeKeyCredentials(before, added);
  await step("Adding the certificate to the registration", () =>
    retryWhileNew(
      fresh,
      sleep,
      () => graph.patch(`applications/${applicationId}`, { keyCredentials: list }),
      (error) => error.status === 404,
    ),
  );
  const after = await step("Checking the registration's certificates", read);
  verifyKeyCredentials(before, after, key);
  return "added";
}

// Gives the key and certificate ACL entries, printing the exact command for whatever could not be
// applied. Returns what is still missing, for the next run to try again.
async function applyAcl(
  state: SetupState,
  acl: AclTools,
  wanted: { privateKey: readonly string[]; certificate: readonly string[] },
  print: (text: string) => void,
): Promise<{ privateKey: string[]; certificate: string[] } | null> {
  const failed = { privateKey: [] as string[], certificate: [] as string[] };
  for (const [which, path] of [
    ["privateKey", state.files.privateKey],
    ["certificate", state.files.certificate],
  ] as const) {
    const named = wanted[which].filter((entry) => !entry.startsWith("mask::"));
    if (named.length === 0) continue;
    const shown = relative(state.cwd, path) || path;
    if (await acl.apply(path, named).catch(() => false))
      print(`Gave ${shown} its access list entries (${named.join(", ")}).`);
    else {
      failed[which] = named;
      print(
        `${shown} needs access list entries the setup could not give it. Run: ${setfaclCommand(path, named)}`,
      );
    }
  }
  return failed.privateKey.length + failed.certificate.length > 0 ? failed : null;
}

// After a rotation's files are in place: the new files get the access list the old ones had, and
// the record names the certificate that was replaced, so later runs can say which one it is and
// try again with any entries that could not be applied.
async function afterRotation(
  state: SetupState,
  plan: SetupPlan,
  acl: AclTools,
  replaced: X509Certificate | null,
  print: (text: string) => void,
) {
  const pendingAcl = await applyAcl(
    state,
    acl,
    {
      privateKey: state.acl.privateKey?.entries ?? [],
      certificate: state.acl.certificate?.entries ?? [],
    },
    print,
  );
  const credential = state.secrets.credential;
  const oldMode = credential.state === "pair" ? credential.privateKeyMode : null;
  if (
    !state.acl.privateKey &&
    oldMode !== null &&
    (oldMode & 0o070) !== 0 &&
    state.platform !== "win32"
  )
    print(
      `The old key could be read by its group or an access list (mode ${oldMode.toString(8)}), and the new one is mode ${plan.writeModes.privateKey.toString(8)}. If the Docker install read it that way, give that access again (docs/SETUP.md).`,
    );
  await writeRecord(state.files, {
    replacedThumbprint: replaced ? thumbprintOf(replaced) : state.record.replacedThumbprint,
    pendingAcl,
  });
}

function finish(
  plan: SetupPlan,
  state: SetupState,
  appId: string | null,
  print: (text: string) => void,
) {
  const lines: string[] = [];
  if (appId)
    lines.push(
      "One step is left, and it stays a click on purpose: an administrator grants the permissions.",
      "Open this link, sign in as an administrator of the tenant, and accept:",
      `  ${consentLink(plan.tenantId, appId)}`,
      "Microsoft then sends the browser to Prism365's address, which may not be running yet. If that page does not load, the consent is still saved.",
      "",
    );
  const defaultSecrets = join(state.cwd, "secrets");
  const prefix =
    state.options.secretsDir === defaultSecrets
      ? ""
      : `PRISM_SECRETS_DIR=${relative(state.cwd, state.options.secretsDir) || state.options.secretsDir} `;
  const envArgument =
    state.options.envFile === join(state.cwd, ".env")
      ? ""
      : ` ${relative(state.cwd, state.options.envFile) || state.options.envFile}`;
  lines.push(
    "Then start Prism365 in Docker:",
    `  ${prefix}scripts/docker-live.sh init${envArgument}`,
    "  scripts/docker-live.sh up",
  );
  lines.push(
    `If an earlier install already made .env.docker, init refuses to replace it: copy ENTRA_CLIENT_ID, ENTRA_ADMIN_OBJECT_IDS and ENABLED_READ_FEATURES from ${relative(state.cwd, state.options.envFile) || ".env"} into it instead, then run up.`,
    "",
    CLI_TOOLS_REMINDER,
    "",
    "Prism365 itself only reads. This setup is the one part of the project that changes anything in Microsoft.",
  );
  print(lines.join("\n"));
}

export async function runSetup(
  options: SetupOptions,
  dependencies: SetupDependencies,
): Promise<SetupOutcome> {
  const now = dependencies.now ?? (() => new Date());
  const sleep =
    dependencies.sleep ?? ((milliseconds) => new Promise((done) => setTimeout(done, milliseconds)));
  const acl = dependencies.acl ?? systemAcl;
  const signedIn = await dependencies.signIn(options.tenant);
  const tenantId = signedIn.tenantId.toLowerCase();
  if (!isUuid(tenantId)) throw new SetupError("Microsoft did not report the tenant ID.");
  if (isUuid(options.tenant) && options.tenant !== tenantId)
    throw new SetupError(
      `You signed in to tenant ${tenantId}, not ${options.tenant}. Run the setup again and sign in with an account of that tenant.`,
    );
  const graph = new SetupGraph({ fetch: dependencies.fetch, getToken: signedIn.getToken, sleep });
  const state = await gatherState(options, graph, tenantId, {
    envExample: dependencies.envExample,
    cwd: dependencies.cwd,
    now: now(),
    acl,
    platform: dependencies.platform ?? process.platform,
  });
  const plan = buildPlan(state);
  dependencies.print(planText(plan, state));
  if (plan.blockers.length > 0)
    throw new SetupError("Nothing was changed. Fix the points above and run the setup again.");
  if (options.dryRun) {
    dependencies.print("\nDry run: nothing was changed.");
    return "dry-run";
  }
  const existingAppId = plan.application.action === "create" ? null : plan.application.appId;
  if (changesNothing(plan)) {
    dependencies.print("\nEverything in the plan is already in place. Nothing was changed.\n");
    finish(plan, state, existingAppId, dependencies.print);
    return "unchanged";
  }
  if (!options.yes && !(await dependencies.confirm("Apply this plan?"))) {
    dependencies.print("Nothing was changed.");
    return "declined";
  }

  // 1. The registration.
  let application: Pick<GraphApplication, "id" | "appId">;
  let marker: string | null = null;
  const fresh = plan.application.action === "create";
  if (plan.application.action === "create") {
    const body = plan.application.body;
    const created = await step("Creating the app registration", () =>
      createApplication(graph, body, state.admin.id, sleep),
    );
    application = created.application;
    marker = created.marker;
    dependencies.print(`Created the app registration (application ID ${application.appId}).`);
  } else {
    application = { id: plan.application.id, appId: plan.application.appId };
    if (plan.application.action === "update") {
      const patch = plan.application.patch;
      await step("Updating the app registration", () =>
        graph.patch(`applications/${application.id}`, patch),
      );
      dependencies.print("Updated the app registration.");
    }
  }

  // 2. .env, as soon as the registration exists: a run that stops after this point is picked up
  // again by the next one, which works on the registration ENTRA_CLIENT_ID names. Read again in
  // case it changed while the plan was on screen.
  const source = await readEnvSource(options, dependencies.envExample);
  const merged = mergeEnv(source.text, managedValues({ ...state, env: source }, application.appId));
  if (!source.exists || merged.text !== source.text) {
    await writeEnvFile(options.envFile, merged.text);
    const keys = merged.changes.map((change) => change.key);
    dependencies.print(
      `${source.exists ? "Updated" : "Created"} ${relative(state.cwd, options.envFile) || options.envFile}${keys.length > 0 ? `: ${keys.join(", ")}` : ""}.`,
    );
  }

  // 3. Its service principal.
  if (plan.servicePrincipal === "create") {
    await step("Creating the service principal", () =>
      createServicePrincipal(graph, application.appId, fresh, sleep),
    );
    dependencies.print("Created the service principal.");
  }

  // 4. Permissions on the secrets, as the plan decided.
  if (plan.restrictDirectory !== null) await setMode(state.files.directory, plan.restrictDirectory);
  if (plan.restrictPrivateKey !== null)
    await setMode(state.files.privateKey, plan.restrictPrivateKey);

  // 5. The certificate. A new key is always saved here before its certificate reaches the
  // registration. A rotation keeps the working pair in place until the new certificate is on
  // the registration, so a failure in between leaves Prism365 signing in as before.
  const files = state.files;
  const credential = state.secrets.credential;
  const request = certificateRequest(options.tenant, now());
  const planned = state.application ? state.keyCredentials : [];
  const upload = async (certificate: X509Certificate) => {
    const result = await addCertificate(graph, application.id, certificate, planned, fresh, sleep);
    dependencies.print(
      result === "added"
        ? "Added the certificate to the registration, and checked that every certificate it had just before is still on it."
        : "The certificate is already on the registration.",
    );
  };
  const shown = (path: string) => relative(state.cwd, path) || path;
  // The certificate a rotation replaces: the working one, unless it is already the new one.
  const replaced = (() => {
    const pending = state.secrets.pending?.certificate;
    const current =
      credential.state === "pair" || credential.state === "mismatch"
        ? credential.certificate
        : null;
    if (current && !(pending && Buffer.compare(pending.raw, current.raw) === 0)) return current;
    return state.secrets.previous;
  })();
  switch (plan.certificate.action) {
    case "create": {
      const made = newCredential(request);
      await writeCredential(files, made, plan.writeModes);
      dependencies.print(`Wrote ${shown(files.privateKey)} and ${shown(files.certificate)}.`);
      await upload(new X509Certificate(made.certificatePem));
      break;
    }
    case "certify": {
      if (credential.state !== "keyOnly")
        throw new SetupError("The private key changed. Run the setup again.");
      const certificatePem = selfSignedCertificate(credential.privateKey, request);
      await writeCredential(files, { privateKeyPem: null, certificatePem }, plan.writeModes);
      dependencies.print(`Wrote ${shown(files.certificate)}.`);
      await upload(new X509Certificate(certificatePem));
      break;
    }
    case "upload":
      if (credential.state === "pair") await upload(credential.certificate);
      break;
    case "rotate": {
      const made = newCredential(request);
      await writePending(files, made, plan.writeModes);
      dependencies.print(
        `Wrote the new pair as ${shown(files.pendingKey)} and ${shown(files.pendingCertificate)}.`,
      );
      await upload(new X509Certificate(made.certificatePem));
      await promotePending(files);
      dependencies.print(
        `Moved the new pair into place. The old one is ${shown(files.previousKey)} and ${shown(files.previousCertificate)}; delete them once you have removed the old certificate from the registration.`,
      );
      await afterRotation(state, plan, acl, replaced, dependencies.print);
      break;
    }
    case "promote":
      await promotePending(files);
      dependencies.print(
        `Finished the earlier rotation. The old pair is ${shown(files.previousKey)} and ${shown(files.previousCertificate)}.`,
      );
      await afterRotation(state, plan, acl, replaced, dependencies.print);
      break;
    case "keep":
    case "blocked":
      break;
  }

  // 6. Access list entries an earlier rotation could not apply.
  if (plan.reapplyAcl) {
    const pendingAcl = await applyAcl(state, acl, plan.reapplyAcl, dependencies.print);
    await writeRecord(files, { ...state.record, pendingAcl });
  }

  // 7. The run labels come off once the run has finished. A first create Microsoft only showed
  // after the second one was made would carry the same marker: if one is there now, it is named.
  if (marker) {
    const label = marker;
    const extras = await findByRunLabel(graph, label, state.admin.id)
      .then((found) => found.filter((item) => item.id !== application.id))
      .catch(() => []);
    for (const extra of extras)
      dependencies.print(
        `Microsoft made a second registration, ${extra.appId}, from a create it had not confirmed. It is not used: delete it in the Entra admin center.`,
      );
    await removeRunLabels(graph, application.id).catch(() =>
      dependencies.print(
        `The registration still carries the label ${label}. It does no harm; the next run takes it off.`,
      ),
    );
  }

  dependencies.print("");
  finish(plan, state, application.appId, dependencies.print);
  return "applied";
}
