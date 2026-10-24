import { createTransport } from "nodemailer";
import { describe, expect, it } from "vitest";
import { destinationCreateSchema } from "@prism365/contracts";
import {
  createSmtpSender,
  type DeliveryConfig,
  deliveryConfigFromEnv,
  type EmailSettings,
  headerText,
  type MailSender,
  recipientAllowed,
  recipientLocalPartAllowed,
  sendEmail,
  smtpErrorCategory,
  smtpConnectionOptions,
} from "../src/delivery-smtp.ts";

// Email alerts go through the operator's own relay, configured by environment only. Every test
// uses a fake sender or nodemailer's in-memory stream transport; none opens a connection.

const files = (content: Record<string, string>) => async (path: string) => {
  const value = content[path];
  if (value === undefined) throw Object.assign(new Error("missing"), { code: "ENOENT" });
  return value;
};
const base = {
  SMTP_URL: "smtps://alerts@relay.example",
  SMTP_PASSWORD_FILE: "/run/secrets/smtp",
  SMTP_FROM: "alerts@example.invalid",
  ALERT_EMAIL_DOMAINS: "example.invalid, Ops.Example.invalid",
};
const load = (env: Record<string, string | undefined>, content = { "/run/secrets/smtp": "pw\n" }) =>
  deliveryConfigFromEnv(env as NodeJS.ProcessEnv, files(content));

describe("relay settings from the environment", () => {
  it("reads smtps with the password from its file", async () => {
    const config = await load(base);
    expect(config.email).toEqual({
      smtp: {
        host: "relay.example",
        port: 465,
        implicitTls: true,
        user: "alerts",
        password: "pw",
        from: "alerts@example.invalid",
      },
      domains: ["example.invalid", "ops.example.invalid"],
    });
  });

  it("takes smtp only with STARTTLS, which is then required", async () => {
    const config = await load({ ...base, SMTP_URL: "smtp://alerts@relay.example" });
    expect(config.email.smtp).toMatchObject({ port: 587, implicitTls: false });
    const options = smtpConnectionOptions(config.email.smtp as NonNullable<EmailSettings["smtp"]>);
    expect(options).toMatchObject({
      secure: false,
      requireTLS: true,
      ignoreTLS: false,
      opportunisticTLS: false,
      tls: { rejectUnauthorized: true, minVersion: "TLSv1.2", servername: "relay.example" },
      logger: false,
      debug: false,
    });
    const implicit = smtpConnectionOptions(
      (await load(base)).email.smtp as NonNullable<EmailSettings["smtp"]>,
    );
    expect(implicit).toMatchObject({ secure: true, tls: { rejectUnauthorized: true } });
  });

  it("refuses any other scheme", async () => {
    for (const SMTP_URL of [
      "http://relay.example",
      "smtp+insecure://relay.example",
      "relay.example",
    ])
      await expect(load({ ...base, SMTP_URL }), SMTP_URL).rejects.toThrow(/SMTP_URL/);
  });

  it("reads the password only from the file", async () => {
    await expect(load({ ...base, SMTP_PASSWORD: "pw" })).rejects.toThrow(
      "SMTP_PASSWORD is never read",
    );
    await expect(load({ ...base, SMTP_URL: "smtps://alerts:pw@relay.example" })).rejects.toThrow(
      "must not carry a password",
    );
    await expect(load({ ...base, SMTP_PASSWORD_FILE: "/missing" })).rejects.toThrow(
      "could not be read",
    );
    await expect(load(base, { "/run/secrets/smtp": "one\ntwo\n" })).rejects.toThrow("one line");
    await expect(load(base, { "/run/secrets/smtp": "\n" })).rejects.toThrow("one line");
    const { SMTP_PASSWORD_FILE: _file, ...withoutFile } = base;
    await expect(load(withoutFile)).rejects.toThrow("SMTP_PASSWORD_FILE is required");
    await expect(load({ ...base, SMTP_URL: "smtps://relay.example" })).rejects.toThrow(
      "must name the user",
    );
  });

  it("takes a relay without a user when it needs no password", async () => {
    const { SMTP_PASSWORD_FILE: _file, ...rest } = base;
    const config = await load({ ...rest, SMTP_URL: "smtps://relay.example:2465" });
    expect(config.email.smtp).toMatchObject({ user: null, password: null, port: 2465 });
    expect(
      smtpConnectionOptions(config.email.smtp as NonNullable<EmailSettings["smtp"]>),
    ).not.toHaveProperty("auth");
  });

  it("refuses options smuggled into the URL", async () => {
    for (const SMTP_URL of [
      "smtps://alerts@relay.example/?tls.rejectUnauthorized=false",
      "smtps://alerts@relay.example/path",
      "smtps://alerts@relay.example#x",
      "smtps://alerts@relay.example\n",
    ])
      await expect(load({ ...base, SMTP_URL }), SMTP_URL).rejects.toThrow(/SMTP_URL/);
  });

  it("needs a sender and recipient domains", async () => {
    await expect(load({ ...base, SMTP_FROM: "Prism <alerts@example.invalid>" })).rejects.toThrow(
      "SMTP_FROM",
    );
    await expect(load({ ...base, ALERT_EMAIL_DOMAINS: "" })).rejects.toThrow("ALERT_EMAIL_DOMAINS");
    await expect(load({ ...base, ALERT_EMAIL_DOMAINS: "*.example.invalid" })).rejects.toThrow(
      "ALERT_EMAIL_DOMAINS",
    );
    await expect(load({ SMTP_FROM: "alerts@example.invalid" })).rejects.toThrow("SMTP_URL");
    expect((await load({})).email).toEqual({ smtp: null, domains: [] });
  });

  it("reads the extra webhook ports", async () => {
    expect((await load({ WEBHOOK_ALLOWED_PORTS: "8443, 9443,8443" })).webhookPorts).toEqual([
      8443, 9443,
    ]);
    for (const WEBHOOK_ALLOWED_PORTS of ["0", "65536", "80a", "-1"])
      await expect(load({ WEBHOOK_ALLOWED_PORTS })).rejects.toThrow("WEBHOOK_ALLOWED_PORTS");
  });
});

describe("sending email", () => {
  const settings = async (): Promise<DeliveryConfig["email"]> => (await load(base)).email;
  function fakeSender() {
    const sent: Record<string, unknown>[] = [];
    const sender: MailSender = {
      sendMail: async (message) => {
        sent.push(message);
        return {};
      },
    };
    return { sent, sender };
  }

  it("allows recipients only in the listed domains, exactly", () => {
    const domains = ["example.invalid"];
    expect(recipientAllowed("ops@example.invalid", domains)).toBe(true);
    expect(recipientAllowed("OPS@Example.Invalid", domains)).toBe(true);
    expect(recipientAllowed("ops@sub.example.invalid", domains)).toBe(false);
    expect(recipientAllowed("ops@example.invalid.attacker.example", domains)).toBe(false);
    expect(recipientAllowed("ops@notexample.invalid", domains)).toBe(false);
    expect(recipientAllowed("ops@example.invalid, x@attacker.example", domains)).toBe(false);
    expect(recipientAllowed('"ops@example.invalid"@attacker.example', domains)).toBe(false);
    expect(recipientAllowed("Ops <ops@example.invalid>", domains)).toBe(false);
  });

  it("refuses local parts a relay could route to another host", () => {
    const domains = ["example.invalid"];
    const refused = [
      "ops%attacker.example@example.invalid",
      "attacker.example!ops@example.invalid",
      '"ops"@example.invalid',
      '"ops@attacker.example"@example.invalid',
      '"a@b"@example.invalid',
      "o..ps@example.invalid",
      ".ops@example.invalid",
      "ops.@example.invalid",
      "ops@attacker.example@example.invalid",
      "ops=x@example.invalid",
      "ops/x@example.invalid",
    ];
    for (const address of refused) {
      expect(recipientAllowed(address, domains), address).toBe(false);
      expect(
        destinationCreateSchema.safeParse({ kind: "email", label: "x", recipients: [address] })
          .success,
        address,
      ).toBe(false);
    }
    for (const local of [
      "ops%attacker.example",
      "attacker.example!ops",
      '"ops"',
      "o..ps",
      ".o",
      "o.",
    ])
      expect(recipientLocalPartAllowed(local), local).toBe(false);
    for (const address of [
      "o'brien@example.invalid",
      "first.last+alerts@example.invalid",
      "a_b-c@example.invalid",
    ]) {
      expect(recipientAllowed(address, domains), address).toBe(true);
      expect(recipientLocalPartAllowed(address.split("@")[0] ?? ""), address).toBe(true);
    }
  });

  it("refuses a routing sender address at start", async () => {
    await expect(load({ ...base, SMTP_FROM: "alerts%relay@example.invalid" })).rejects.toThrow(
      "SMTP_FROM",
    );
  });

  it("checks every recipient again at send time and sends nothing when one is outside", async () => {
    const { sent, sender } = fakeSender();
    const outcome = await sendEmail(await settings(), sender, {
      recipients: ["ops@example.invalid", "someone@attacker.example"],
      subject: "Synthetic",
      text: "Synthetic",
    });
    expect(outcome).toEqual({ ok: false, error: "recipient_not_allowed" });
    expect(sent).toEqual([]);
  });

  it("refuses to send without a relay", async () => {
    const { sent, sender } = fakeSender();
    expect(
      await sendEmail({ smtp: null, domains: [] }, sender, {
        recipients: ["ops@example.invalid"],
        subject: "x",
        text: "x",
      }),
    ).toEqual({ ok: false, error: "email_not_configured" });
    expect(sent).toEqual([]);
  });

  it("passes recipients as mailboxes and envelope, never as a string to parse", async () => {
    const { sent, sender } = fakeSender();
    expect(
      await sendEmail(await settings(), sender, {
        recipients: ["ops@example.invalid"],
        subject: "Alert\r\nBcc: someone@attacker.example",
        text: "Body",
      }),
    ).toEqual({ ok: true });
    expect(sent[0]).toMatchObject({
      from: { name: "Prism365", address: "alerts@example.invalid" },
      to: [{ name: "", address: "ops@example.invalid" }],
      envelope: { from: "alerts@example.invalid", to: ["ops@example.invalid"] },
      subject: "Alert Bcc: someone@attacker.example",
      disableFileAccess: true,
      disableUrlAccess: true,
    });
  });

  it("never lets a subject or file name start a header of its own", async () => {
    for (const evil of ["a\r\nBcc: x@attacker.example", "a\nX: y", "a\u2028X: y", "a\u0085X: y"])
      expect(headerText(evil)).not.toMatch(/[\r\n\u2028\u0085]/);
    // Composed by nodemailer itself, in memory.
    const transport = createTransport({ streamTransport: true, buffer: true });
    const raw: string[] = [];
    const composing: MailSender = {
      sendMail: async (message) => {
        const info = await transport.sendMail(message);
        raw.push(String(info.message));
        return info;
      },
    };
    await sendEmail(await settings(), composing, {
      recipients: ["ops@example.invalid"],
      subject: "Alert\r\nBcc: someone@attacker.example",
      text: "Body",
      attachment: {
        filename: "audit\r\nContent-Type: text/html.csv",
        content: Buffer.from("a,b\n1,2\n"),
      },
    });
    const headers = (raw[0] ?? "").split("\r\n\r\n")[0] ?? "";
    expect(headers).not.toMatch(/^Bcc:/im);
    // The would-be header stays inside the one Subject line.
    expect(headers).toMatch(/^Subject: Alert Bcc: someone@attacker\.example$/m);
    expect(raw[0]).not.toMatch(/^Content-Type: text\/html/im);
    expect(headers).toMatch(/^To: ops@example\.invalid$/m);
    expect(headers).toMatch(/^Auto-Submitted: auto-generated$/m);
  });

  it("refuses an attachment taken from a path or a URL", async () => {
    const transport = createTransport({ streamTransport: true, buffer: true });
    const composing: MailSender = { sendMail: (message) => transport.sendMail(message) };
    const smtp = (await settings()).smtp as NonNullable<EmailSettings["smtp"]>;
    // The message options sendEmail sets turn file and URL access off even if a path slipped in.
    await expect(
      composing.sendMail({
        from: smtp.from,
        to: "ops@example.invalid",
        text: "x",
        disableFileAccess: true,
        disableUrlAccess: true,
        attachments: [{ path: "/etc/hosts" }],
      }),
    ).rejects.toMatchObject({ code: "EFILEACCESS" });
  });

  it("never sends twice when the relay goes quiet after taking the message", async () => {
    const smtp = await settings();
    const attempt = async (behaviour: "after-data" | "before-data" | "answered" | "reset") => {
      const controller = new AbortController();
      let seen: AbortSignal | null = null;
      const sender: MailSender = {
        sendMail: (_message, control) => {
          seen = control?.signal ?? null;
          if (behaviour !== "before-data") control?.dataStarted();
          if (behaviour === "answered")
            return Promise.reject(Object.assign(new Error("try later"), { responseCode: 451 }));
          if (behaviour === "reset")
            return Promise.reject(Object.assign(new Error("closed"), { code: "ECONNECTION" }));
          return new Promise(() => {});
        },
      };
      setTimeout(() => controller.abort(), 20);
      const outcome = await sendEmail(
        smtp,
        sender,
        { recipients: ["ops@example.invalid"], subject: "s", text: "t" },
        controller.signal,
      );
      return { outcome, closed: (seen as AbortSignal | null)?.aborted ?? false };
    };
    // Cut off once the message was being handed over: it may have arrived, so it is not retried.
    expect(await attempt("after-data")).toEqual({
      outcome: { ok: false, error: "outcome_unknown" },
      closed: true,
    });
    expect((await attempt("reset")).outcome).toEqual({ ok: false, error: "outcome_unknown" });
    // Before that, it is an ordinary timeout, and the connection is closed too.
    expect(await attempt("before-data")).toEqual({
      outcome: { ok: false, error: "timeout" },
      closed: true,
    });
    // A refusal the relay spells out is what it says.
    expect((await attempt("answered")).outcome).toEqual({ ok: false, error: "server_error" });
  });

  it("opens no connection for an attempt already ended", async () => {
    const smtp = (await settings()).smtp as NonNullable<EmailSettings["smtp"]>;
    const controller = new AbortController();
    controller.abort();
    let started = false;
    await expect(
      createSmtpSender(smtp).sendMail(
        {
          from: smtp.from,
          to: "ops@example.invalid",
          text: "x",
          envelope: { from: smtp.from, to: ["ops@example.invalid"] },
        },
        {
          signal: controller.signal,
          dataStarted: () => {
            started = true;
          },
        },
      ),
    ).rejects.toMatchObject({ name: "AbortError" });
    expect(started).toBe(false);
  });

  it("maps relay errors to fixed categories", () => {
    expect(smtpErrorCategory({ code: "EAUTH" })).toBe("smtp_auth_failed");
    expect(smtpErrorCategory({ code: "ETLS" })).toBe("tls_failed");
    expect(smtpErrorCategory({ code: "ECONNECTION" })).toBe("connection_failed");
    expect(smtpErrorCategory({ code: "ETIMEDOUT" })).toBe("timeout");
    expect(smtpErrorCategory({ code: "EENVELOPE", responseCode: 550 })).toBe("smtp_rejected");
    expect(smtpErrorCategory({ code: "EENVELOPE", responseCode: 451 })).toBe("server_error");
    expect(smtpErrorCategory(new Error("x"))).toBe("error");
  });
});
