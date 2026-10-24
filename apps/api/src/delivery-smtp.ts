import { readFile } from "node:fs/promises";
import { isIP } from "node:net";
import { type DeliveryError, emailAddressSchema } from "@prism365/contracts";
import { Readable } from "node:stream";
import MailComposer from "nodemailer/lib/mail-composer";
import SMTPConnection, { type SMTPConnectionOptions } from "nodemailer/lib/smtp-connection";

// Email alerts through the operator's own mail relay (docs/alerts.md, ADR-0011). Never Graph
// mail: Prism365 holds no permission to send as anyone, and asking for one would break its
// read-only promise. The relay is configured by environment only, so no request and no page can
// point mail somewhere else or read the password:
//
//   SMTP_URL             smtps://relay.example:465 (TLS from the start) or
//                        smtp://relay.example:587 (STARTTLS, which is then required); a user
//                        may be named in it, a password never
//   SMTP_PASSWORD_FILE   a file holding the password, read once at start
//   SMTP_FROM            the sender address
//   ALERT_EMAIL_DOMAINS  the only domains recipients may be in, comma-separated
//
// Plain SMTP without TLS is refused, and so is a certificate that does not verify.

export interface SmtpSettings {
  host: string;
  port: number;
  // TLS from the first byte (smtps). Otherwise STARTTLS is required before anything is sent.
  implicitTls: boolean;
  user: string | null;
  password: string | null;
  from: string;
}

export interface EmailSettings {
  smtp: SmtpSettings | null;
  domains: string[];
}

export interface DeliveryConfig {
  email: EmailSettings;
  // Ports other than 443 a webhook may use (WEBHOOK_ALLOWED_PORTS).
  webhookPorts: number[];
}

const domainSchema = /^(?!-)[a-z0-9-]{1,63}(?<!-)(?:\.(?!-)[a-z0-9-]{1,63}(?<!-))+$/;
const unsafeText = /[\0\r\n\u0085\u2028\u2029]/;

export async function deliveryConfigFromEnv(
  env: NodeJS.ProcessEnv,
  read: (path: string) => Promise<string> = (path) => readFile(path, "utf8"),
): Promise<DeliveryConfig> {
  if (env.SMTP_PASSWORD !== undefined)
    throw new Error(
      "SMTP_PASSWORD is never read. Put the relay password in a file and set SMTP_PASSWORD_FILE.",
    );
  const webhookPorts = [
    ...new Set(
      (env.WEBHOOK_ALLOWED_PORTS ?? "")
        .split(",")
        .map((value) => value.trim())
        .filter(Boolean)
        .map((value) => {
          const port = /^[0-9]{1,5}$/.test(value) ? Number(value) : Number.NaN;
          if (!(port >= 1 && port <= 65_535))
            throw new Error("WEBHOOK_ALLOWED_PORTS must list port numbers from 1 to 65535.");
          return port;
        }),
    ),
  ];
  if (webhookPorts.length > 20) throw new Error("WEBHOOK_ALLOWED_PORTS lists at most 20 ports.");
  if (!env.SMTP_URL) {
    if (env.SMTP_FROM || env.SMTP_PASSWORD_FILE || env.ALERT_EMAIL_DOMAINS)
      throw new Error("Email alerts need SMTP_URL as well as SMTP_FROM and ALERT_EMAIL_DOMAINS.");
    return { email: { smtp: null, domains: [] }, webhookPorts };
  }
  const raw = env.SMTP_URL;
  let url: URL;
  try {
    if (/[\s\p{Cc}\\]/u.test(raw)) throw new Error("whitespace");
    url = new URL(raw);
  } catch {
    throw new Error("SMTP_URL must be a URL such as smtps://relay.example:465.");
  }
  if (url.protocol !== "smtps:" && url.protocol !== "smtp:")
    throw new Error("SMTP_URL must start with smtps:// or smtp:// (which then requires STARTTLS).");
  if (url.password) throw new Error("SMTP_URL must not carry a password. Use SMTP_PASSWORD_FILE.");
  if ((url.pathname !== "" && url.pathname !== "/") || url.search || url.hash)
    throw new Error("SMTP_URL takes a host, a port and optionally a user, nothing else.");
  const host = url.hostname.replace(/^\[|\]$/g, "");
  if (!host) throw new Error("SMTP_URL needs a host.");
  const implicitTls = url.protocol === "smtps:";
  const port = url.port ? Number(url.port) : implicitTls ? 465 : 587;
  let user: string | null = null;
  if (url.username) {
    try {
      user = decodeURIComponent(url.username);
    } catch {
      throw new Error("SMTP_URL has a user name that cannot be read.");
    }
    if (unsafeText.test(user)) throw new Error("SMTP_URL has a user name that cannot be read.");
  }
  let password: string | null = null;
  if (env.SMTP_PASSWORD_FILE) {
    let content: string;
    try {
      content = await read(env.SMTP_PASSWORD_FILE);
    } catch {
      throw new Error("SMTP_PASSWORD_FILE could not be read.");
    }
    password = content.replace(/\r?\n$/, "");
    if (password.length === 0 || password.length > 512 || unsafeText.test(password))
      throw new Error("SMTP_PASSWORD_FILE must hold the password on one line.");
  }
  if (user && password === null)
    throw new Error("SMTP_URL names a user, so SMTP_PASSWORD_FILE is required.");
  if (!user && password !== null)
    throw new Error("SMTP_PASSWORD_FILE is set, so SMTP_URL must name the user.");
  const from = emailAddressSchema.safeParse(env.SMTP_FROM ?? "");
  if (!from.success)
    throw new Error("SMTP_FROM must be one plain address, such as alerts@example.invalid.");
  const domains = [
    ...new Set(
      (env.ALERT_EMAIL_DOMAINS ?? "")
        .split(",")
        .map((value) => value.trim().toLowerCase())
        .filter(Boolean),
    ),
  ];
  if (domains.length === 0 || domains.length > 50 || !domains.every((d) => domainSchema.test(d)))
    throw new Error(
      "ALERT_EMAIL_DOMAINS must list up to 50 domain names, such as example.invalid.",
    );
  return {
    email: { smtp: { host, port, implicitTls, user, password, from: from.data }, domains },
    webhookPorts,
  };
}

// A recipient is allowed only in one of the listed domains exactly; a subdomain is not the
// domain. The address must be one plain mailbox, as the destination schema requires, and its
// local part is checked here again on its own, so an address a relay could route elsewhere
// (ops%attacker.example@, attacker.example!ops@, a quoted local part) never reaches the relay
// even if the schema ever widened.
const localPart = /^[A-Za-z0-9_+'-]+(?:\.[A-Za-z0-9_+'-]+)*$/;

export function recipientLocalPartAllowed(local: string): boolean {
  return local.length > 0 && local.length <= 64 && localPart.test(local);
}

export function recipientAllowed(address: string, domains: readonly string[]): boolean {
  const parsed = emailAddressSchema.safeParse(address);
  if (!parsed.success) return false;
  const parts = parsed.data.split("@");
  if (parts.length !== 2) return false;
  const [local = "", domain = ""] = parts;
  if (!recipientLocalPartAllowed(local)) return false;
  return domains.includes(domain);
}

// One line of header text: no line break of any kind and no other control character, so a
// subject or a file name taken from provider text can never start a header of its own.
export function headerText(value: string, max = 200): string {
  return value
    .replace(/[\p{Cc}\p{Zl}\p{Zp}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, max);
}

export interface EmailMessage {
  recipients: readonly string[];
  subject: string;
  text: string;
  attachment?: { filename: string; content: Buffer };
}

// What an attempt tells the sender, and learns from it: the signal aborts it (the connection is
// closed at once), and dataStarted is called when the relay asks for the message itself, after
// which a lost connection may still have delivered it.
export interface SendControl {
  signal: AbortSignal;
  dataStarted(): void;
}

export interface MailSender {
  sendMail(message: Record<string, unknown>, control?: SendControl): Promise<unknown>;
}

// The connection options for the relay. TLS is never optional: smtps starts with it, smtp must
// upgrade with STARTTLS or the message is not sent. Certificates are verified.
export function smtpConnectionOptions(smtp: SmtpSettings): SMTPConnectionOptions {
  return {
    host: smtp.host,
    port: smtp.port,
    secure: smtp.implicitTls,
    requireTLS: !smtp.implicitTls,
    ignoreTLS: false,
    opportunisticTLS: false,
    // The certificate is checked against the host either way; a server name for SNI must not be
    // an address.
    tls: {
      rejectUnauthorized: true,
      minVersion: "TLSv1.2",
      ...(isIP(smtp.host) === 0 ? { servername: smtp.host } : {}),
    },
    connectionTimeout: 10_000,
    greetingTimeout: 10_000,
    socketTimeout: 20_000,
    logger: false,
    debug: false,
  };
}

// The message, read only when the relay has answered DATA: nothing reads a stream before that,
// so the first read is the moment the message starts to leave.
function messageStream(raw: Buffer, control: SendControl | undefined): Readable {
  let sent = false;
  return new Readable({
    read() {
      if (sent) return;
      sent = true;
      control?.dataStarted();
      this.push(raw);
      this.push(null);
    },
  });
}

// One connection per message, on nodemailer's SMTP connection, so the attempt owns it: an abort
// closes it, and the DATA phase is visible. The message is composed in memory first, with file
// and URL access off, so an attachment can only be the content handed over.
export function createSmtpSender(smtp: SmtpSettings): MailSender {
  return {
    async sendMail(message, control) {
      const raw = await new MailComposer({
        ...message,
        disableFileAccess: true,
        disableUrlAccess: true,
      })
        .compile()
        .build();
      const envelope = message.envelope as { from: string; to: string[] };
      return new Promise((resolve, reject) => {
        const connection = new SMTPConnection(smtpConnectionOptions(smtp));
        let settled = false;
        const finish = (error: unknown, info?: unknown) => {
          if (settled) return;
          settled = true;
          control?.signal.removeEventListener("abort", abort);
          if (error) {
            connection.close();
            reject(error);
          } else {
            connection.quit();
            resolve(info);
          }
        };
        const abort = () => finish(Object.assign(new Error("aborted"), { name: "AbortError" }));
        if (control?.signal.aborted) return abort();
        control?.signal.addEventListener("abort", abort, { once: true });
        connection.on("error", (error: unknown) => finish(error));
        connection.connect(() => {
          const send = () =>
            connection.send(
              { from: envelope.from, to: envelope.to, size: raw.length },
              messageStream(raw, control),
              (error, info) => finish(error, info),
            );
          if (smtp.user && smtp.password !== null)
            connection.login({ user: smtp.user, pass: smtp.password }, (error) =>
              error ? finish(error) : send(),
            );
          else send();
        });
      });
    },
  };
}

export function smtpErrorCategory(error: unknown): DeliveryError {
  const code =
    typeof error === "object" && error !== null && "code" in error ? String(error.code) : "";
  const responseCode =
    typeof error === "object" && error !== null && "responseCode" in error
      ? Number(error.responseCode)
      : 0;
  if (code === "EAUTH" || code === "ENOAUTH") return "smtp_auth_failed";
  if (code === "ETLS" || code === "EREQUIRETLS") return "tls_failed";
  if (code === "ETIMEDOUT") return "timeout";
  if (code === "ECONNECTION" || code === "ESOCKET" || code === "EDNS" || code === "EPROXY")
    return "connection_failed";
  if (responseCode >= 400 && responseCode < 500) return "server_error";
  if (code === "EENVELOPE" || code === "EMESSAGE" || responseCode >= 500) return "smtp_rejected";
  return "error";
}

// Sends one message. Every recipient is checked against ALERT_EMAIL_DOMAINS again at send time,
// because the setting may have changed since the destination was saved; one recipient outside
// them refuses the whole message. Recipients go as separate mailbox objects and as the envelope,
// so nothing is ever parsed out of a string.
//
// `signal` ends the attempt. A connection lost without an answer once the relay asked for the
// message (a timeout, a reset, the abort) is outcome_unknown: the message may have arrived, so it
// is never tried again. An explicit answer from the relay always says what happened.
export async function sendEmail(
  settings: EmailSettings,
  sender: MailSender | null,
  message: EmailMessage,
  signal: AbortSignal = new AbortController().signal,
): Promise<{ ok: true } | { ok: false; error: DeliveryError }> {
  const smtp = settings.smtp;
  if (!smtp || !sender) return { ok: false, error: "email_not_configured" };
  if (
    message.recipients.length === 0 ||
    !message.recipients.every((address) => recipientAllowed(address, settings.domains))
  )
    return { ok: false, error: "recipient_not_allowed" };
  let dataStarted = false;
  const control: SendControl = {
    signal,
    dataStarted: () => {
      dataStarted = true;
    },
  };
  const aborted = new Promise<never>((_resolve, reject) => {
    const stop = () => reject(Object.assign(new Error("aborted"), { name: "AbortError" }));
    if (signal.aborted) stop();
    signal.addEventListener("abort", stop, { once: true });
  });
  aborted.catch(() => {});
  try {
    await Promise.race([
      sender.sendMail(
        {
          from: { name: "Prism365", address: smtp.from },
          to: message.recipients.map((address) => ({ name: "", address })),
          envelope: { from: smtp.from, to: [...message.recipients] },
          subject: headerText(message.subject),
          text: message.text,
          headers: { "Auto-Submitted": "auto-generated" },
          disableFileAccess: true,
          disableUrlAccess: true,
          ...(message.attachment
            ? {
                attachments: [
                  {
                    filename: headerText(message.attachment.filename, 120),
                    content: message.attachment.content,
                    contentType: "text/csv; charset=utf-8",
                  },
                ],
              }
            : {}),
        },
        control,
      ),
      aborted,
    ]);
    return { ok: true };
  } catch (error) {
    const answered =
      typeof error === "object" && error !== null && "responseCode" in error
        ? Number(error.responseCode) > 0
        : false;
    const category = signal.aborted ? "timeout" : smtpErrorCategory(error);
    if (dataStarted && !answered && !["smtp_auth_failed", "tls_failed"].includes(category))
      return { ok: false, error: "outcome_unknown" };
    return { ok: false, error: category };
  }
}
