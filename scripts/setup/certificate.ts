import {
  createPrivateKey,
  createPublicKey,
  generateKeyPairSync,
  type KeyObject,
  randomBytes,
  sign,
  X509Certificate,
} from "node:crypto";
import { chmod, link, lstat, mkdir, open, readFile, rename, stat, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";

// The certificate Prism365 signs in to Microsoft with: an RSA 2048 key and a self-signed X.509
// certificate, built here with Node's own crypto so the setup needs no extra tool. RSA, because
// MSAL signs its client assertion with PS256 when it is given a SHA-256 thumbprint.

// Microsoft accepts a certificate credential for at most one year from its start, so the
// certificate lasts 365 days: never longer than a year, even across a leap day.
export const CERTIFICATE_DAYS = 365;
// An X.509 common name holds at most 64 characters.
const MAX_COMMON_NAME = 64;

function derLength(length: number): Buffer {
  if (length < 0x80) return Buffer.from([length]);
  const bytes: number[] = [];
  for (let rest = length; rest > 0; rest = Math.floor(rest / 256)) bytes.unshift(rest % 256);
  return Buffer.from([0x80 | bytes.length, ...bytes]);
}

function der(tag: number, ...content: Buffer[]): Buffer {
  const body = Buffer.concat(content);
  return Buffer.concat([Buffer.from([tag]), derLength(body.length), body]);
}

const sequence = (...content: Buffer[]) => der(0x30, ...content);

function objectIdentifier(dotted: string): Buffer {
  const [first = 0, second = 0, ...rest] = dotted.split(".").map(Number);
  const bytes = [first * 40 + second];
  for (const arc of rest) {
    const chunk = [arc % 128];
    for (let value = Math.floor(arc / 128); value > 0; value = Math.floor(value / 128))
      chunk.unshift(0x80 | (value % 128));
    bytes.push(...chunk);
  }
  return der(0x06, Buffer.from(bytes));
}

// A non-negative INTEGER in its shortest form: no redundant leading zero, and a zero in front
// when the first bit would otherwise make it negative.
function positiveInteger(value: Buffer): Buffer {
  let start = 0;
  while (start < value.length - 1 && value[start] === 0) start += 1;
  const trimmed = value.subarray(start);
  return der(0x02, (trimmed[0] ?? 0) & 0x80 ? Buffer.concat([Buffer.from([0]), trimmed]) : trimmed);
}

// UTCTime up to 2049 and GeneralizedTime from 2050, as RFC 5280 requires.
function certificateTime(date: Date): Buffer {
  const digits = date.toISOString().slice(0, 19).replace(/[-:T]/g, "");
  const year = date.getUTCFullYear();
  return year < 2050
    ? der(0x17, Buffer.from(`${digits.slice(2)}Z`, "ascii"))
    : der(0x18, Buffer.from(`${digits}Z`, "ascii"));
}

function distinguishedName(commonName: string): Buffer {
  return sequence(
    der(0x31, sequence(objectIdentifier("2.5.4.3"), der(0x0c, Buffer.from(commonName, "utf8")))),
  );
}

const sha256WithRsa = sequence(objectIdentifier("1.2.840.113549.1.1.11"), Buffer.from([5, 0]));

function extensions(): Buffer {
  const critical = Buffer.from([0x01, 0x01, 0xff]);
  // Not a certificate authority, and the key only signs (the client assertion).
  const basicConstraints = sequence(objectIdentifier("2.5.29.19"), critical, der(0x04, sequence()));
  const keyUsage = sequence(
    objectIdentifier("2.5.29.15"),
    critical,
    der(0x04, der(0x03, Buffer.from([0x07, 0x80]))),
  );
  return der(0xa3, sequence(basicConstraints, keyUsage));
}

export function commonNameFor(tenant: string): string {
  return `Prism365 ${tenant}`.slice(0, MAX_COMMON_NAME);
}

export interface CertificateRequest {
  commonName: string;
  notBefore: Date;
  notAfter: Date;
}

// A self-signed certificate for an existing private key.
export function selfSignedCertificate(privateKey: KeyObject, request: CertificateRequest): string {
  if (privateKey.asymmetricKeyType !== "rsa")
    throw new Error("The private key must be an RSA key.");
  if (request.notAfter.getTime() <= request.notBefore.getTime())
    throw new Error("The certificate must end after it starts.");
  // RFC 5280 allows up to 20 bytes. Sixteen random bytes with the top bit clear and the next one
  // set stay positive and never need a padding byte.
  const serial = randomBytes(16);
  serial[0] = ((serial[0] ?? 0) & 0x7f) | 0x40;
  const name = distinguishedName(request.commonName);
  const subjectPublicKeyInfo = createPublicKey(privateKey).export({ type: "spki", format: "der" });
  const tbsCertificate = sequence(
    der(0xa0, positiveInteger(Buffer.from([2]))),
    positiveInteger(serial),
    sha256WithRsa,
    name,
    sequence(certificateTime(request.notBefore), certificateTime(request.notAfter)),
    name,
    subjectPublicKeyInfo,
    extensions(),
  );
  const signature = sign("sha256", tbsCertificate, privateKey);
  const certificate = sequence(
    tbsCertificate,
    sha256WithRsa,
    der(0x03, Buffer.from([0]), signature),
  );
  const lines = certificate.toString("base64").match(/.{1,64}/g) ?? [];
  return `-----BEGIN CERTIFICATE-----\n${lines.join("\n")}\n-----END CERTIFICATE-----\n`;
}

export interface NewCredential {
  privateKeyPem: string;
  certificatePem: string;
}

export function certificateRequest(tenant: string, now: Date): CertificateRequest {
  const notBefore = new Date(Math.floor(now.getTime() / 1_000) * 1_000);
  return {
    commonName: commonNameFor(tenant),
    notBefore,
    notAfter: new Date(notBefore.getTime() + CERTIFICATE_DAYS * 86_400_000),
  };
}

export function newCredential(request: CertificateRequest): NewCredential {
  const { privateKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
  return {
    privateKeyPem: privateKey.export({ type: "pkcs8", format: "pem" }).toString(),
    certificatePem: selfSignedCertificate(privateKey, request),
  };
}

// What Graph's keyCredential `key` holds: the certificate's DER bytes in base64.
export function certificateKey(certificate: X509Certificate): string {
  return certificate.raw.toString("base64");
}

export interface CertificateFiles {
  directory: string;
  privateKey: string;
  certificate: string;
}

export function certificateFiles(directory: string): CertificateFiles {
  return {
    directory,
    privateKey: join(directory, "entra-private.pem"),
    certificate: join(directory, "entra-public.pem"),
  };
}

// The pair already in the secrets directory, as far as the setup needs to know it. The private
// key is read only to check that it belongs to the certificate; it is never printed or sent.
export type LocalCredential =
  | { state: "none" }
  | { state: "pair"; certificate: X509Certificate; privateKeyMode: number | null }
  | { state: "keyOnly"; privateKey: KeyObject; privateKeyMode: number | null }
  | { state: "certificateOnly" }
  | { state: "mismatch" }
  | { state: "unreadable"; file: "certificate" | "privateKey" };

export interface LocalSecrets {
  credential: LocalCredential;
  // Permission bits of the secrets directory, or null when it does not exist yet.
  directoryMode: number | null;
}

async function modeOf(path: string): Promise<number | null> {
  try {
    return (await stat(path)).mode & 0o777;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw error;
  }
}

async function readIfPresent(path: string): Promise<Buffer | null> {
  try {
    return await readFile(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw error;
  }
}

export async function inspectSecrets(files: CertificateFiles): Promise<LocalSecrets> {
  const directoryMode = await modeOf(files.directory);
  const keyBytes = await readIfPresent(files.privateKey);
  const certificateBytes = await readIfPresent(files.certificate);
  let privateKey: KeyObject | null = null;
  if (keyBytes) {
    try {
      privateKey = createPrivateKey(keyBytes);
    } catch {
      return { directoryMode, credential: { state: "unreadable", file: "privateKey" } };
    }
  }
  let certificate: X509Certificate | null = null;
  if (certificateBytes) {
    try {
      certificate = new X509Certificate(certificateBytes);
    } catch {
      return { directoryMode, credential: { state: "unreadable", file: "certificate" } };
    }
  }
  const privateKeyMode = privateKey ? await modeOf(files.privateKey) : null;
  if (privateKey && certificate)
    return {
      directoryMode,
      credential: certificate.checkPrivateKey(privateKey)
        ? { state: "pair", certificate, privateKeyMode }
        : { state: "mismatch" },
    };
  if (privateKey)
    return { directoryMode, credential: { state: "keyOnly", privateKey, privateKeyMode } };
  if (certificate) return { directoryMode, credential: { state: "certificateOnly" } };
  return { directoryMode, credential: { state: "none" } };
}

const ownerOnly = (mode: number | null) => mode === null || (mode & 0o077) === 0;
export const restrictsToOwner = ownerOnly;

async function temporaryFile(path: string, content: string, mode: number): Promise<string> {
  const temporary = join(dirname(path), `.${randomBytes(6).toString("hex")}.tmp`);
  const handle = await open(temporary, "wx", mode);
  try {
    await handle.writeFile(content);
    await handle.sync();
  } finally {
    await handle.close();
  }
  // The mode given to open is reduced by the umask; set it exactly.
  if (process.platform !== "win32") await chmod(temporary, mode);
  return temporary;
}

async function removeQuietly(path: string) {
  await unlink(path).catch(() => {});
}

// Puts a finished temporary file in place. Without `replace` an existing file is never
// overwritten: the hard link fails if the name is taken.
async function place(temporary: string, path: string, replace: boolean) {
  if (replace) {
    await rename(temporary, path);
    return;
  }
  try {
    await link(temporary, path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "EEXIST")
      throw new Error(`${path} already exists. Pass --rotate-certificate to replace it.`);
    // A file system without hard links: check, then rename.
    const existing = await lstat(path).catch(() => null);
    if (existing)
      throw new Error(`${path} already exists. Pass --rotate-certificate to replace it.`);
    await rename(temporary, path);
    return;
  }
  await removeQuietly(temporary);
}

export async function prepareSecretsDirectory(directory: string) {
  await mkdir(directory, { recursive: true, mode: 0o700 });
  if (process.platform !== "win32" && !ownerOnly(await modeOf(directory)))
    await chmod(directory, 0o700);
}

// Writes the key (0600) and certificate (0644). Both are written to temporary files first and
// the key is put in place before the certificate, so an interruption leaves at worst a key
// without its certificate, which a re-run certifies again. The key is never overwritten unless
// `replaceKey` (--rotate-certificate).
export async function writeCredential(
  files: CertificateFiles,
  credential: { privateKeyPem: string | null; certificatePem: string },
  options: { replaceKey: boolean },
) {
  await prepareSecretsDirectory(files.directory);
  const key = credential.privateKeyPem
    ? await temporaryFile(files.privateKey, credential.privateKeyPem, 0o600)
    : null;
  let certificate: string | null = null;
  try {
    certificate = await temporaryFile(files.certificate, credential.certificatePem, 0o644);
    if (key) await place(key, files.privateKey, options.replaceKey);
    await place(certificate, files.certificate, true);
  } finally {
    if (key) await removeQuietly(key);
    if (certificate) await removeQuietly(certificate);
  }
}

// Tightens a private key file that is readable by others.
export async function restrictPrivateKey(files: CertificateFiles) {
  if (process.platform !== "win32") await chmod(files.privateKey, 0o600);
}
