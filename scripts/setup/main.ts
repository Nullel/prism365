import { stdin, stdout } from "node:process";
import { createInterface } from "node:readline/promises";
import { fileURLToPath } from "node:url";
import { PublicClientApplication } from "@azure/msal-node";
import { terminalText } from "./graph.ts";
import { parseOptions, UsageError, usage } from "./options.ts";
import { runSetup, type SignedIn } from "./setup.ts";

// `npm run setup -- --tenant <tenant>`: see docs/SETUP.md. Signs in with a device code, shows
// its plan, and applies it after confirmation.

// Microsoft's own public client for the Microsoft Graph command line tools, which supports the
// device code flow, so no registration has to exist before this one is made.
const GRAPH_COMMAND_LINE_TOOLS = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const SCOPES = [
  "https://graph.microsoft.com/Application.ReadWrite.All",
  "https://graph.microsoft.com/User.Read",
];

// MSAL keeps its token cache in memory only (no cache plugin), so nothing about the sign-in is
// written to disk, and the token never leaves this process except to Microsoft Graph.
async function signIn(tenant: string): Promise<SignedIn> {
  const client = new PublicClientApplication({
    auth: {
      clientId: GRAPH_COMMAND_LINE_TOOLS,
      authority: `https://login.microsoftonline.com/${tenant}`,
    },
    system: { loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } },
  });
  console.log(
    "Sign in with an account that can create app registrations (Application Administrator, Cloud Application Administrator or Global Administrator).",
  );
  const result = await client.acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: (response) => console.log(terminalText(response.message, 500)),
  });
  const account = result?.account;
  if (!result || !account) throw new Error("Microsoft did not complete the sign-in.");
  console.log("");
  return {
    tenantId: result.tenantId,
    getToken: async () =>
      (await client.acquireTokenSilent({ account, scopes: SCOPES })).accessToken,
  };
}

async function confirm(question: string): Promise<boolean> {
  if (!stdin.isTTY) {
    console.log(`${question} This terminal cannot answer; run it again with --yes to apply.`);
    return false;
  }
  const prompt = createInterface({ input: stdin, output: stdout });
  try {
    const answer = await prompt.question(`\n${question} Type yes to apply it: `);
    return /^(y|yes)$/i.test(answer.trim());
  } finally {
    prompt.close();
  }
}

// MSAL errors carry an error code and, in the message, Microsoft's AADSTS code; only those are
// shown, never the rest of the provider's message.
function signInFailure(error: unknown): string | null {
  if (!error || typeof error !== "object" || !("errorCode" in error)) return null;
  const code = String((error as { errorCode: unknown }).errorCode).slice(0, 80);
  const message = error instanceof Error ? error.message : "";
  const aadsts = /AADSTS\d{4,8}/.exec(message)?.[0];
  const reason =
    code === "authorization_declined"
      ? "The sign-in was declined."
      : code === "expired_token" || code === "code_expired"
        ? "The code expired before the sign-in finished. Run the setup again."
        : aadsts === "AADSTS65001"
          ? "Microsoft needs consent for the Microsoft Graph command line tools to use Application.ReadWrite.All. Sign in as an administrator who can grant it."
          : "Microsoft did not complete the sign-in.";
  return `${reason} (${[code, aadsts].filter(Boolean).join(", ")})`;
}

async function main() {
  let options: ReturnType<typeof parseOptions>;
  try {
    options = parseOptions(process.argv.slice(2), process.cwd());
  } catch (error) {
    if (!(error instanceof UsageError)) throw error;
    console.error(`${error.message}\n\n${usage}`);
    process.exitCode = 2;
    return;
  }
  if (options === "help") {
    console.log(usage);
    return;
  }
  try {
    await runSetup(options, {
      fetch: (input, init) => fetch(input, init),
      signIn,
      confirm,
      print: (text) => console.log(text),
      envExample: fileURLToPath(new URL("../../.env.example", import.meta.url)),
      cwd: process.cwd(),
    });
  } catch (error) {
    console.error(
      `\n${signInFailure(error) ?? (error instanceof Error ? error.message : "The setup failed.")}`,
    );
    process.exitCode = 1;
  }
}

await main();
