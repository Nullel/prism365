import { describe, expect, it } from "vitest";
import {
  type AlertMessage,
  maskAddresses,
  minimise,
  renderEmail,
  renderWebhook,
  slackText,
  teamsText,
} from "../src/delivery-payload.ts";

// What leaves the server: the minimum by default, details only when asked for, and full IP
// addresses only in email with details on. Synthetic names and documentation addresses only.

const origin = "https://prism365.example";
const meta = {
  deliveryId: "11111111-2222-4333-8444-555555555555",
  sentAt: "2026-10-24T12:00:00.000Z",
};

function message(overrides: Partial<AlertMessage> = {}): AlertMessage {
  return {
    kind: "finding",
    title: "A new or escalated finding at high severity",
    severity: "high",
    count: 1,
    items: [
      {
        title: "Password spray pattern from one address",
        severity: "high",
        subject: "203.0.113.45",
        note: "new",
        details: {
          summary: "203.0.113.45 failed to sign in to 12 accounts (2001:db8:85a3::8a2e:370:7334).",
          network: "203.0.113.0/24",
          actorPrincipal: "avery.stone@example.invalid",
          attempts: 40,
        },
      },
    ],
    lines: ["Rule: Spray from 198.51.100.7"],
    link: "/#/signins",
    ...overrides,
  };
}

const everything = (value: unknown) => JSON.stringify(value);

describe("payload minimisation", () => {
  it("sends title, severity, counts, subject and link by default, and no details", () => {
    for (const kind of ["teams", "slack", "generic", "email"] as const) {
      const minimal = minimise(message(), { kind, includeDetails: false });
      expect(minimal.items[0]?.details, kind).toEqual({});
      expect(everything(minimal), kind).not.toContain("avery.stone");
      expect(everything(minimal), kind).not.toContain("failed to sign in");
    }
  });

  it("never sends a full IP address to a webhook, details or not", () => {
    for (const kind of ["teams", "slack", "generic"] as const)
      for (const includeDetails of [false, true]) {
        const minimal = minimise(message(), { kind, includeDetails });
        const sent = renderWebhook(kind, minimal, origin, meta);
        expect(sent, `${kind} ${includeDetails}`).not.toContain("203.0.113.45");
        expect(sent).not.toContain("198.51.100.7");
        expect(sent).not.toContain("8a2e:370:7334");
        expect(sent).toContain("203.0.113.0/24");
      }
  });

  it("keeps full addresses only for email with details on", () => {
    const plain = renderEmail(
      minimise(message(), { kind: "email", includeDetails: false }),
      origin,
    );
    expect(`${plain.subject}${plain.text}`).not.toContain("203.0.113.45");
    expect(plain.text).toContain("203.0.113.0/24");
    const detailed = renderEmail(
      minimise(message(), { kind: "email", includeDetails: true }),
      origin,
    );
    expect(detailed.text).toContain("203.0.113.45");
    expect(detailed.text).toContain("2001:db8:85a3::8a2e:370:7334");
    expect(detailed.text).toContain("avery.stone@example.invalid");
  });

  it("adds details for a webhook that asked for them, with addresses cut", () => {
    const minimal = minimise(message(), { kind: "generic", includeDetails: true });
    expect(minimal.items[0]?.details).toMatchObject({
      actorPrincipal: "avery.stone@example.invalid",
      attempts: 40,
    });
    expect(minimal.items[0]?.details.summary).toBe(
      "203.0.113.0/24 failed to sign in to 12 accounts (2001:db8:85a3::/48).",
    );
  });

  it("finds addresses in every common spelling", () => {
    const cases: [string, string][] = [
      ["from 203.0.113.45.", "from 203.0.113.0/24."],
      ["203.0.113.45:443", "203.0.113.0/24:443"],
      ["[2001:db8::1]:443", "[2001:db8:0::/48]:443"],
      ["2001:DB8::ABCD", "2001:db8:0::/48"],
      ["at ::1 and ::ffff:c000:201", "at 0:0:0::/48 and ::ffff:192.0.2.0/24"],
      ["::ffff:192.0.2.1", "::ffff:192.0.2.0/24"],
      ["203.0.113.45/32", "203.0.113.0/24/32"],
      ["203.0.113.0/24", "203.0.113.0/24"],
      ["fe80::1%eth0", "fe80:0:0::/48"],
      ["address:2001:db8::7", "address:2001:db8:0::/48"],
      ["010.000.002.001", "10.0.2.0/24"],
    ];
    for (const [input, output] of cases) expect(maskAddresses(input), input).toBe(output);
    // A full address with a port or a hex prefix glued on is not one address, but it still holds
    // one: the whole run is cut.
    for (const glued of [
      "2001:db8:1:2:3:4:5:6:443",
      "de:2001:db8:1:2:3:4:5:6",
      "id ab:2001:db8:1:2:3:4:5:6",
      "ab:2001:db8::5:6:7:8",
      "2001:db8:1:2:3:4:5:6::",
    ]) {
      const masked = maskAddresses(glued);
      expect(masked, glued).not.toContain("3:4:5:6");
      expect(masked, glued).not.toContain("5:6:7:8");
      expect(masked, glued).toContain("::/48");
    }
    // An IPv4 address with a digit or a dotted part glued on still holds the address, so the
    // run is cut too, and so is a percent-encoded IPv6 address.
    for (const [glued, address] of [
      ["967.178.80.106", "67.178.80.106"],
      ["72.183.241.749", "72.183.241.74"],
      ["116.120.148.233.7", "116.120.148.233"],
      ["1.128.200.224.66", "128.200.224.66"],
      ["version 1.2.3.4.5", "1.2.3.4"],
      ["ip=2001%3Adb8%3A85a3%3A%3A8a2e", "2001:db8:85a3::8a2e"],
      ["ip=203%2E0%2E113%2E45", "203.0.113.45"],
    ] as const) {
      const masked = maskAddresses(glued);
      expect(masked, glued).not.toContain(address);
      expect(masked, glued).not.toMatch(/%3a|%2e/i);
    }
    // Times, versions and ordinary text are left alone.
    for (const text of [
      "12:30:45",
      "version 1.2.3",
      "build 10.0.19045.3803",
      "pi is 3.14159",
      "a: b: c",
      "aa:bb:cc:dd:ee:ff",
      "10:00:00 to 11:30:00",
    ])
      expect(maskAddresses(text), text).toBe(text);
  });

  it("keeps at most ten items and twelve detail fields", () => {
    const many = message({
      count: 30,
      items: Array.from({ length: 30 }, (_, index) => ({
        title: `Finding ${index}`,
        severity: "low" as const,
        subject: null,
        note: null,
        details: Object.fromEntries(Array.from({ length: 20 }, (_, field) => [`f${field}`, field])),
      })),
    });
    const minimal = minimise(many, { kind: "generic", includeDetails: true });
    expect(minimal.items).toHaveLength(10);
    expect(Object.keys(minimal.items[0]?.details ?? {})).toHaveLength(12);
    expect(renderEmail(minimal, origin).text).toContain("And 20 more in Prism365.");
  });
});

describe("webhook formats", () => {
  it("sends Teams one Adaptive Card of text blocks with one link back", () => {
    const sent = JSON.parse(
      renderWebhook(
        "teams",
        minimise(message(), { kind: "teams", includeDetails: false }),
        origin,
        meta,
      ),
    );
    expect(sent.type).toBe("message");
    expect(sent.attachments).toHaveLength(1);
    const card = sent.attachments[0];
    expect(card.contentType).toBe("application/vnd.microsoft.card.adaptive");
    expect(card.content.type).toBe("AdaptiveCard");
    expect(card.content.body.length).toBeGreaterThan(0);
    for (const block of card.content.body) expect(block.type).toBe("TextBlock");
    expect(card.content.actions).toEqual([
      { type: "Action.OpenUrl", title: "Open Prism365", url: "https://prism365.example/#/signins" },
    ]);
  });

  it("breaks links and mentions that provider text tries to make", () => {
    const evil = message({
      items: [
        {
          title: "[Open Prism365](https://phish.example/login)",
          severity: "high",
          subject: "<!channel> <https://phish.example|Click> @here",
          note: null,
          details: {},
        },
      ],
    });
    for (const kind of ["teams", "slack"] as const) {
      const sent = renderWebhook(
        kind,
        minimise(evil, { kind, includeDetails: false }),
        origin,
        meta,
      );
      expect(sent, kind).not.toContain("](https://");
      expect(sent, kind).not.toContain("https://phish.example");
    }
    const slack = JSON.parse(
      renderWebhook(
        "slack",
        minimise(evil, { kind: "slack", includeDetails: false }),
        origin,
        meta,
      ),
    );
    expect(slack.text).not.toContain("<!channel>");
    expect(slack.text).not.toContain("<https://phish");
    expect(slack.text).toContain("&lt;!channel&gt;");
    expect(slack.text).not.toMatch(/@here\b/);
    // The one real link is Prism365's own.
    expect(slack.text.match(/<https?:[^>]*>/g)).toEqual([
      "<https://prism365.example/#/signins|Open Prism365>",
    ]);
    expect(slack).toMatchObject({ unfurl_links: false, unfurl_media: false });
    expect(slackText("a & b")).toBe("a &amp; b");
    expect(teamsText("see https://phish.example")).toBe("see https[:]//phish.example");
  });

  it("documents the generic format as version 1", () => {
    const sent = JSON.parse(
      renderWebhook(
        "generic",
        minimise(message(), { kind: "generic", includeDetails: false }),
        origin,
        meta,
      ),
    );
    expect(sent).toMatchObject({
      version: 1,
      type: "prism365.alert",
      deliveryId: meta.deliveryId,
      kind: "finding",
      severity: "high",
      count: 1,
      link: "https://prism365.example/#/signins",
    });
    expect(Object.keys(sent).sort()).toEqual([
      "count",
      "deliveryId",
      "items",
      "kind",
      "lines",
      "link",
      "sentAt",
      "severity",
      "title",
      "type",
      "version",
    ]);
  });

  it("keeps provider line breaks out of every line", () => {
    const broken = message({ title: "Line one\r\nLine two\u2028three" });
    const email = renderEmail(minimise(broken, { kind: "email", includeDetails: false }), origin);
    expect(email.subject).toBe("Prism365: Line one Line two three");
    expect(email.text.split("\n")[0]).toBe("Line one Line two three");
  });
});
