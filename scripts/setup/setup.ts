import { X509Certificate } from "node:crypto";
import { readFile } from "node:fs/promises";
import { join, relative } from "node:path";
import { isUuid } from "../../apps/api/src/config.ts";
import {
  type CertificateFiles,
  certificateFiles,
  certificateKey,
  certificateRequest,
  inspectSecrets,
  newCredential,
  prepareSecretsDirectory,
  restrictPrivateKey,
  selfSignedCertificate,
  writeCredential,
} from "./certificate.ts";
import { mergeEnv, writeEnvFile } from "./env-file.ts";
import {
  applicationListSchema,
  applicationSchema,
  explainGraphError,
  type Fetch,
  type GraphApplication,
  graphServicePrincipalSchema,
  keyCredentialsResponseSchema,
  meSchema,
  SetupGraph,
  SetupGraphError,
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
  SETUP_TAG,
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
}

export type SetupOutcome = "applied" | "unchanged" | "dry-run" | "declined";

// A failure the administrator can act on, already in plain words.
export class SetupError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SetupError";
  }
}

const APPLICATION_FIELDS = "id,appId,displayName,signInAudience,tags,web,requiredResourceAccess";
// A registration made moments ago can take a little while to be visible everywhere in Microsoft
// Entra. Steps on it wait and retry, about 45 seconds in all.
const NEW_OBJECT_DELAYS = [2_000, 4_000, 8_000, 15_000, 15_000];

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

async function findApplication(graph: SetupGraph, options: SetupOptions) {
  const select = `$select=${APPLICATION_FIELDS}`;
  if (options.appId) {
    try {
      const application = await graph.get(
        `applications(appId='${options.appId}')?${select}`,
        applicationSchema,
      );
      return { application, foundBy: "appId" as const, matches: 1 };
    } catch (error) {
      if (error instanceof SetupGraphError && error.status === 404)
        return { application: null, foundBy: null, matches: 0 };
      throw error;
    }
  }
  const search = async (filter: string) => {
    const page = await graph.get(
      `applications?$filter=${encodeURIComponent(filter)}&${select}&$top=10`,
      applicationListSchema,
    );
    return { items: page.value, count: page.value.length + (page["@odata.nextLink"] ? 1 : 0) };
  };
  const tagged = await search(`tags/any(t:t eq '${SETUP_TAG}')`);
  if (tagged.count > 0)
    return {
      application: tagged.count === 1 ? (tagged.items[0] ?? null) : null,
      foundBy: "tag" as const,
      matches: tagged.count,
    };
  const named = await search(`displayName eq '${options.name.replaceAll("'", "''")}'`);
  return {
    application: named.count === 1 ? (named.items[0] ?? null) : null,
    foundBy: named.count > 0 ? ("name" as const) : null,
    matches: named.count,
  };
}

export async function gatherState(
  options: SetupOptions,
  graph: SetupGraph,
  tenantId: string,
  dependencies: Pick<SetupDependencies, "envExample" | "cwd"> & { now: Date },
): Promise<SetupState> {
  const admin = await step("Reading the signed-in account", () =>
    graph.get("me?$select=id,userPrincipalName", meSchema),
  );
  const found = await step("Looking for an existing Prism365 registration", () =>
    findApplication(graph, options),
  );
  const graphPrincipal = await step("Reading Microsoft Graph's permission list", () =>
    graph.get(
      `servicePrincipals(appId='${GRAPH_RESOURCE_APP_ID}')?$select=id,oauth2PermissionScopes,appRoles`,
      graphServicePrincipalSchema,
    ),
  );
  let servicePrincipal = false;
  let keyCredentials: unknown[] = [];
  const application = found.application;
  if (application) {
    servicePrincipal = await step("Checking the registration's service principal", async () => {
      const list = await graph.get(
        `servicePrincipals?$filter=${encodeURIComponent(`appId eq '${application.appId}'`)}&$select=id`,
        servicePrincipalListSchema,
      );
      return list.value.length > 0;
    });
    keyCredentials = await step("Reading the registration's certificates", async () => {
      const response = await graph.get(
        `applications/${application.id}?$select=keyCredentials`,
        keyCredentialsResponseSchema,
      );
      return response.keyCredentials;
    });
  }
  const files = certificateFiles(options.secretsDir);
  return {
    options,
    tenantId,
    admin: { id: admin.id, userPrincipalName: admin.userPrincipalName ?? null },
    application,
    foundBy: found.foundBy,
    matches: found.matches,
    servicePrincipal,
    graph: graphPrincipal,
    keyCredentials,
    files,
    secrets: await inspectSecrets(files),
    env: await readEnvSource(options, dependencies.envExample),
    cwd: dependencies.cwd,
    now: dependencies.now,
  };
}

function changesNothing(plan: SetupPlan) {
  return (
    plan.application.action === "keep" &&
    plan.servicePrincipal === "keep" &&
    plan.certificate.action === "keep" &&
    !plan.restrictDirectory &&
    !plan.restrictPrivateKey &&
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

// Adds the certificate to the registration without touching any other: the current list is
// read with its keys, sent back unchanged with the new one, and read again to check.
async function addCertificate(
  graph: SetupGraph,
  applicationId: string,
  certificate: X509Certificate,
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

async function applyCertificate(
  plan: SetupPlan,
  state: SetupState,
  now: Date,
): Promise<X509Certificate | null> {
  const files: CertificateFiles = state.files;
  const credential = state.secrets.credential;
  const request = certificateRequest(state.options.tenant, now);
  switch (plan.certificate.action) {
    case "create":
    case "rotate": {
      const made = newCredential(request);
      await writeCredential(files, made, { replaceKey: plan.certificate.action === "rotate" });
      return new X509Certificate(made.certificatePem);
    }
    case "certify": {
      if (credential.state !== "keyOnly")
        throw new SetupError("The private key changed. Run the setup again.");
      const certificatePem = selfSignedCertificate(credential.privateKey, request);
      await writeCredential(files, { privateKeyPem: null, certificatePem }, { replaceKey: false });
      return new X509Certificate(certificatePem);
    }
    case "upload":
      return credential.state === "pair" ? credential.certificate : null;
    case "keep":
    case "blocked":
      return null;
  }
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
  const fresh = plan.application.action === "create";
  if (plan.application.action === "create") {
    const body = plan.application.body;
    application = await step("Creating the app registration", () =>
      graph.post("applications", body, applicationSchema),
    );
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

  // 2. Its service principal.
  if (plan.servicePrincipal === "create") {
    await step("Creating the service principal", () =>
      retryWhileNew(
        fresh,
        sleep,
        () =>
          graph.post(
            "servicePrincipals",
            { appId: application.appId, tags: ["WindowsAzureActiveDirectoryIntegratedApp"] },
            servicePrincipalSchema,
          ),
        (error) => error.status === 404 || error.status === 400,
      ),
    );
    dependencies.print("Created the service principal.");
  }

  // 3. The certificate, written before it is uploaded so a key never reaches the registration
  // without being saved here first.
  if (plan.restrictDirectory) await prepareSecretsDirectory(state.files.directory);
  if (plan.restrictPrivateKey) await restrictPrivateKey(state.files);
  const certificate = await applyCertificate(plan, state, now());
  if (plan.certificate.action === "create" || plan.certificate.action === "rotate")
    dependencies.print(
      `Wrote ${relative(state.cwd, state.files.privateKey)} and ${relative(state.cwd, state.files.certificate)}.`,
    );
  else if (plan.certificate.action === "certify")
    dependencies.print(`Wrote ${relative(state.cwd, state.files.certificate)}.`);

  // 4. The certificate on the registration.
  if (certificate) {
    const result = await addCertificate(graph, application.id, certificate, fresh, sleep);
    dependencies.print(
      result === "added"
        ? "Added the certificate to the registration. Every certificate that was there before is still there."
        : "The certificate is already on the registration.",
    );
  }

  // 5. .env, read again in case it changed while the plan was on screen.
  const source = await readEnvSource(options, dependencies.envExample);
  const merged = mergeEnv(source.text, managedValues({ ...state, env: source }, application.appId));
  if (!source.exists || merged.text !== source.text) {
    await writeEnvFile(options.envFile, merged.text);
    const keys = merged.changes.map((change) => change.key);
    dependencies.print(
      `${source.exists ? "Updated" : "Created"} ${relative(state.cwd, options.envFile) || options.envFile}${keys.length > 0 ? `: ${keys.join(", ")}` : ""}.`,
    );
  }
  dependencies.print("");
  finish(plan, state, application.appId, dependencies.print);
  return "applied";
}
