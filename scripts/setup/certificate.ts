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
  // A rotation's new pair until it is on the registration, and the pair it replaced afterwards.
  // They end in .pem like the pair itself, so .gitignore, .dockerignore and the Dockerfile's guard
  // keep them out of Git and out of images too.
  pendingKey: string;
  pendingCertificate: string;
  previousKey: string;
  previousCertificate: string;
  // What the setup remembers between runs (see SetupRecord). No key material.
  record: string;
}

export function certificateFiles(directory: string): CertificateFiles {
  return {
    directory,
    privateKey: join(directory, "entra-private.pem"),
    certificate: join(directory, "entra-public.pem"),
    pendingKey: join(directory, "entra-private.new.pem"),
    pendingCertificate: join(directory, "entra-public.new.pem"),
    previousKey: join(directory, "entra-private.previous.pem"),
    previousCertificate: join(directory, "entra-public.previous.pem"),
    record: join(directory, "setup-state.json"),
  };
}

// MSAL signs its client assertion with PS256, so the key must be RSA, and Microsoft wants at
// least 2048 bits. Anything else would only fail later, after Microsoft had been changed.
export function keyProblem(privateKey: KeyObject): string | null {
  if (privateKey.asymmetricKeyType !== "rsa")
    return `is ${privateKey.asymmetricKeyType ? `an ${privateKey.asymmetricKeyType.toUpperCase()}` : "not an RSA"} key, and Prism365 signs in with RSA only`;
  const bits = privateKey.asymmetricKeyDetails?.modulusLength ?? 0;
  return bits < 2048 ? `is a ${bits}-bit RSA key, and Microsoft needs at least 2048 bits` : null;
}

// The pair already in the secrets directory, as far as the setup needs to know it. The private
// key is read only to check its type and that it belongs to the certificate; it is never printed
// or sent.
export type LocalCredential =
  | { state: "none" }
  | {
      state: "pair";
      certificate: X509Certificate;
      privateKey: KeyObject;
      privateKeyMode: number | null;
    }
  | { state: "keyOnly"; privateKey: KeyObject; privateKeyMode: number | null }
  | { state: "certificateOnly" }
  | { state: "mismatch"; certificate: X509Certificate; privateKey: KeyObject }
  | { state: "unreadable"; file: "certificate" | "privateKey" };

export interface LocalSecrets {
  credential: LocalCredential;
  // Permission bits of the secrets directory, or null when it does not exist yet.
  directoryMode: number | null;
  // A rotation's new certificate, with its key while the key has not been moved into place yet
  // (null once it has); `pendingFiles` is true when either file is there at all.
  pending: { certificate: X509Certificate; privateKey: KeyObject | null } | null;
  pendingFiles: boolean;
  // The pair the last rotation replaced: its certificate, and whether either file is there.
  previous: X509Certificate | null;
  previousFiles: boolean;
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

const parseKey = (bytes: Buffer | null) => {
  if (!bytes) return null;
  try {
    return createPrivateKey(bytes);
  } catch {
    return "unreadable" as const;
  }
};
const parseCertificate = (bytes: Buffer | null) => {
  if (!bytes) return null;
  try {
    return new X509Certificate(bytes);
  } catch {
    return "unreadable" as const;
  }
};

async function readPair(keyPath: string, certificatePath: string) {
  const keyBytes = await readIfPresent(keyPath);
  const certificateBytes = await readIfPresent(certificatePath);
  return {
    present: keyBytes !== null || certificateBytes !== null,
    privateKey: parseKey(keyBytes),
    certificate: parseCertificate(certificateBytes),
  };
}

export async function inspectSecrets(files: CertificateFiles): Promise<LocalSecrets> {
  const directoryMode = await modeOf(files.directory);
  const current = await readPair(files.privateKey, files.certificate);
  const pendingPair = await readPair(files.pendingKey, files.pendingCertificate);
  const previousPair = await readPair(files.previousKey, files.previousCertificate);
  const pendingCertificate =
    pendingPair.certificate && pendingPair.certificate !== "unreadable"
      ? pendingPair.certificate
      : null;
  const pendingKey =
    pendingPair.privateKey && pendingPair.privateKey !== "unreadable"
      ? pendingPair.privateKey
      : null;
  // With both files there they must belong together; a lone certificate is what an interrupted
  // move leaves once its key is in place.
  const pending =
    pendingCertificate &&
    (pendingKey ? pendingCertificate.checkPrivateKey(pendingKey) : pendingPair.privateKey === null)
      ? { certificate: pendingCertificate, privateKey: pendingKey }
      : null;
  const around = {
    directoryMode,
    pending,
    pendingFiles: pendingPair.present,
    previous:
      previousPair.certificate && previousPair.certificate !== "unreadable"
        ? previousPair.certificate
        : null,
    previousFiles: previousPair.present,
  };
  const { privateKey, certificate } = current;
  if (privateKey === "unreadable")
    return { ...around, credential: { state: "unreadable", file: "privateKey" } };
  if (certificate === "unreadable")
    return { ...around, credential: { state: "unreadable", file: "certificate" } };
  const privateKeyMode = privateKey ? await modeOf(files.privateKey) : null;
  if (privateKey && certificate)
    return {
      ...around,
      credential: certificate.checkPrivateKey(privateKey)
        ? { state: "pair", certificate, privateKey, privateKeyMode }
        : { state: "mismatch", certificate, privateKey },
    };
  if (privateKey)
    return { ...around, credential: { state: "keyOnly", privateKey, privateKeyMode } };
  if (certificate) return { ...around, credential: { state: "certificateOnly" } };
  return { ...around, credential: { state: "none" } };
}

// True when a rotation was interrupted while its files were being moved into place: the
// current key and certificate do not match, and one of them is already the new pair's.
export function interruptedRotation(secrets: LocalSecrets): boolean {
  const { credential, pending } = secrets;
  if (credential.state !== "mismatch" || !pending) return false;
  return (
    pending.certificate.checkPrivateKey(credential.privateKey) ||
    pending.certificate.raw.equals(credential.certificate.raw)
  );
}

async function temporaryFile(path: string, content: string, mode: number): Promise<string> {
  const temporary = join(dirname(path), `.${randomBytes(6).toString("hex")}.tmp`);
  const handle = await open(temporary, "wx", mode);
  try {
    await handle.writeFile(content);
    await handle.sync();
  } finally {
    await handle.close();
  }
  // The mode given to open is reduced by the umask; set it exactly. Windows has no such modes:
  // there the folder's access control list decides.
  if (process.platform !== "win32") await chmod(temporary, mode);
  return temporary;
}

async function removeQuietly(path: string) {
  await unlink(path).catch(() => {});
}

const exists = (path: string) =>
  lstat(path).then(
    () => true,
    () => false,
  );

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
    if (await exists(path))
      throw new Error(`${path} already exists. Pass --rotate-certificate to replace it.`);
    await rename(temporary, path);
    return;
  }
  await removeQuietly(temporary);
}

// Creates the secrets directory, owner-only, when it is missing. An existing directory's
// permissions are changed only by an explicit step of the plan (see plan.ts), because they may
// carry an access control list the administrator set on purpose.
export async function ensureSecretsDirectory(directory: string) {
  await mkdir(directory, { recursive: true, mode: 0o700 });
}

export interface WriteModes {
  // 0600 normally; 0640 when the directory has a default ACL, so the entries it hands down are
  // not masked off (with an ACL the group bits are the mask, not the owning group).
  privateKey: number;
  certificate: number;
}
export const ownerModes: WriteModes = { privateKey: 0o600, certificate: 0o644 };

// Writes a first pair, or a certificate for a key that is already there. Both are written to
// temporary files first and the key is put in place before the certificate, so an interruption
// leaves at worst a key without its certificate, which a re-run certifies again. An existing key
// is never overwritten: replacing a working pair is a rotation (writePending, then
// promotePending).
export async function writeCredential(
  files: CertificateFiles,
  credential: { privateKeyPem: string | null; certificatePem: string },
  modes: WriteModes = ownerModes,
) {
  await ensureSecretsDirectory(files.directory);
  const key = credential.privateKeyPem
    ? await temporaryFile(files.privateKey, credential.privateKeyPem, modes.privateKey)
    : null;
  let certificate: string | null = null;
  try {
    certificate = await temporaryFile(
      files.certificate,
      credential.certificatePem,
      modes.certificate,
    );
    if (key) await place(key, files.privateKey, false);
    await place(certificate, files.certificate, true);
  } finally {
    if (key) await removeQuietly(key);
    if (certificate) await removeQuietly(certificate);
  }
}

// A rotation's new pair, written beside the working one under .new names. The working pair is
// not touched, so if the upload fails Prism365 keeps signing in with it.
export async function writePending(
  files: CertificateFiles,
  credential: NewCredential,
  modes: WriteModes = ownerModes,
) {
  await ensureSecretsDirectory(files.directory);
  const key = await temporaryFile(files.pendingKey, credential.privateKeyPem, modes.privateKey);
  let certificate: string | null = null;
  try {
    certificate = await temporaryFile(
      files.pendingCertificate,
      credential.certificatePem,
      modes.certificate,
    );
    await rename(key, files.pendingKey);
    await rename(certificate, files.pendingCertificate);
  } finally {
    await removeQuietly(key);
    if (certificate) await removeQuietly(certificate);
  }
}

// Moves a rotation's new pair into place once it is on the registration, keeping the replaced
// pair under .previous names until the administrator removes it. Each file is first linked to
// its .previous name and then replaced in one rename, so the working name always holds a file.
// It can be run again after an interruption and carries on where it stopped; an interruption
// between the key and the certificate leaves a mismatched pair that interruptedRotation spots.
export async function promotePending(files: CertificateFiles) {
  for (const [pending, current, previous] of [
    [files.pendingKey, files.privateKey, files.previousKey],
    [files.pendingCertificate, files.certificate, files.previousCertificate],
  ] as const) {
    if (!(await exists(pending))) continue;
    if ((await exists(current)) && !(await exists(previous))) await link(current, previous);
    await rename(pending, current);
  }
}

// What the setup remembers between runs, in setup-state.json beside the key (mode 0600). It
// holds a thumbprint, a flag and an application ID, never key material and never ACL entries:
// anyone who can write in the folder could edit it, so it is only ever a hint.
export interface SetupRecord {
  // The SHA-256 thumbprint of the certificate the last rotation replaced, so a later run can say
  // which certificate on the registration it is once the .previous files are gone.
  replacedThumbprint: string | null;
  // A rotation could not copy the old files' access list to the new ones. The entries themselves
  // are read again from the .previous files, never taken from here.
  aclPending: boolean;
  // The registration this install used before --app-id moved it to another one.
  previousClientId: string | null;
}

export const emptyRecord: SetupRecord = {
  replacedThumbprint: null,
  aclPending: false,
  previousClientId: null,
};
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

export async function readRecord(files: CertificateFiles): Promise<SetupRecord> {
  const bytes = await readIfPresent(files.record);
  if (!bytes) return emptyRecord;
  try {
    const value = JSON.parse(bytes.toString("utf8")) as Record<string, unknown>;
    return {
      replacedThumbprint:
        typeof value.replacedThumbprint === "string" &&
        /^[0-9A-F]{64}$/.test(value.replacedThumbprint)
          ? value.replacedThumbprint
          : null,
      aclPending: value.aclPending === true,
      previousClientId:
        typeof value.previousClientId === "string" && GUID.test(value.previousClientId)
          ? value.previousClientId
          : null,
    };
  } catch {
    return emptyRecord;
  }
}

export async function writeRecord(files: CertificateFiles, record: SetupRecord) {
  await ensureSecretsDirectory(files.directory);
  const temporary = await temporaryFile(
    files.record,
    `${JSON.stringify(record, null, 2)}\n`,
    0o600,
  );
  try {
    await rename(temporary, files.record);
  } finally {
    await removeQuietly(temporary);
  }
}

// The thumbprint Graph and the Entra admin center use for a certificate, without colons.
export const thumbprintOf = (certificate: X509Certificate) =>
  certificate.fingerprint256.replaceAll(":", "");

// Sets a file or directory's permission bits, when the plan decided to.
export async function setMode(path: string, mode: number) {
  if (process.platform !== "win32") await chmod(path, mode);
}
