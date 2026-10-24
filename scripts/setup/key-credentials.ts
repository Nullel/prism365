import { z } from "zod";

// A PATCH of `keyCredentials` replaces the whole list, so adding Prism365's certificate means
// sending every existing certificate back exactly as Microsoft returned it. These checks make
// sure that can never drop another credential: if any existing key cannot be read in full the
// setup stops, and after the update every key ID from before must still be there.

const keyCredentialSchema = z
  .object({
    keyId: z.string().regex(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i),
    key: z.string().nullable().optional(),
  })
  .loose();

export const keyCredentialListSchema = z.array(z.unknown()).max(500);

export interface NewKeyCredential {
  type: "AsymmetricX509Cert";
  usage: "Verify";
  key: string;
  displayName: string;
  startDateTime: string;
  endDateTime: string;
}

export class KeyCredentialSafetyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "KeyCredentialSafetyError";
  }
}

function decoded(key: string | null | undefined): Buffer | null {
  if (typeof key !== "string" || !/^[A-Za-z0-9+/]+={0,2}$/.test(key.replace(/\s/g, "")))
    return null;
  const bytes = Buffer.from(key, "base64");
  return bytes.length > 0 ? bytes : null;
}

function parsed(list: readonly unknown[]) {
  return list.map((item) => {
    const credential = keyCredentialSchema.safeParse(item);
    if (!credential.success)
      throw new KeyCredentialSafetyError(
        "Microsoft returned a certificate credential without a readable key ID, so the setup stopped rather than risk removing it. Nothing was changed on the registration.",
      );
    return credential.data;
  });
}

// True when the list already holds this certificate (compared as bytes, not as text).
export function hasCertificate(list: readonly unknown[], key: string): boolean {
  const wanted = decoded(key);
  if (!wanted) return false;
  return parsed(list).some((credential) => decoded(credential.key)?.equals(wanted) ?? false);
}

// Key IDs of credentials whose key value Microsoft did not return in full.
export function unreadableKeys(list: readonly unknown[]): string[] {
  return parsed(list)
    .filter((credential) => decoded(credential.key) === null)
    .map((credential) => credential.keyId);
}

// The list to PATCH: every existing entry unchanged, then the new certificate.
export function mergeKeyCredentials(
  existing: readonly unknown[],
  added: NewKeyCredential,
): unknown[] {
  if (!decoded(added.key)) throw new KeyCredentialSafetyError("The new certificate is empty.");
  const missing = unreadableKeys(existing);
  if (missing.length > 0)
    throw new KeyCredentialSafetyError(
      `Microsoft did not return the full key of ${missing.length === 1 ? "certificate" : "certificates"} ${missing.join(", ")} on this registration. Sending the list back without it would delete it, so the setup stopped. Nothing was changed on the registration. Add the certificate in the Entra admin center instead (docs/SETUP.md).`,
    );
  return [...existing, added];
}

// Run on the list read back after the PATCH.
export function verifyKeyCredentials(
  before: readonly unknown[],
  after: readonly unknown[],
  addedKey: string,
): void {
  const now = parsed(after);
  const lost = parsed(before)
    .map((credential) => credential.keyId.toLowerCase())
    .filter((keyId) => !now.some((credential) => credential.keyId.toLowerCase() === keyId));
  if (lost.length > 0)
    throw new KeyCredentialSafetyError(
      `After the update the registration no longer lists ${lost.length === 1 ? "certificate" : "certificates"} ${lost.join(", ")} that ${lost.length === 1 ? "was" : "were"} there before. Check Certificates and secrets on the registration in the Entra admin center now and add any missing certificate back.`,
    );
  if (!hasCertificate(after, addedKey))
    throw new KeyCredentialSafetyError(
      "Microsoft accepted the update, but the new certificate is not on the registration. Run the setup again.",
    );
}
