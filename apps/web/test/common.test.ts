import { readFileSync } from "node:fs";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import type { Finding } from "@prism365/contracts";
import {
  day,
  FindingItem,
  RelativeTime,
  relativeTime,
  SeverityDot,
  SkeletonRows,
} from "../src/common.tsx";
import { at, relativeCell, severityDot } from "./markup.ts";

const NOW = Date.parse("2026-10-24T12:00:00Z");
const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;
const before = (milliseconds: number) => new Date(NOW - milliseconds).toISOString();

describe("relativeTime", () => {
  it("says just now for the first minute, including a clock a little ahead", () => {
    expect(relativeTime(before(0), NOW)).toBe("just now");
    expect(relativeTime(before(59_999), NOW)).toBe("just now");
    expect(relativeTime(before(-4 * MINUTE), NOW)).toBe("just now");
  });

  it("counts whole minutes and hours, rounding down", () => {
    expect(relativeTime(before(MINUTE), NOW)).toBe("1 min ago");
    expect(relativeTime(before(5 * MINUTE), NOW)).toBe("5 min ago");
    expect(relativeTime(before(HOUR - 1), NOW)).toBe("59 min ago");
    expect(relativeTime(before(HOUR), NOW)).toBe("1 h ago");
    expect(relativeTime(before(3 * HOUR + 40 * MINUTE), NOW)).toBe("3 h ago");
    expect(relativeTime(before(DAY - 1), NOW)).toBe("23 h ago");
  });

  it("says yesterday for the day after the first, then counts days", () => {
    expect(relativeTime(before(DAY), NOW)).toBe("yesterday");
    expect(relativeTime(before(2 * DAY - 1), NOW)).toBe("yesterday");
    expect(relativeTime(before(2 * DAY), NOW)).toBe("2 days ago");
    expect(relativeTime(before(4 * DAY + 5 * HOUR), NOW)).toBe("4 days ago");
    expect(relativeTime(before(7 * DAY - 1), NOW)).toBe("6 days ago");
  });

  it("gives the date from a week on, and for a time well in the future", () => {
    const week = before(7 * DAY);
    expect(relativeTime(week, NOW)).toBe(day(week));
    expect(relativeTime(week, NOW)).not.toMatch(/ago/);
    const ahead = before(-6 * MINUTE);
    expect(relativeTime(ahead, NOW)).toBe(day(ahead));
  });

  it("does not guess when there is no time or it cannot be read", () => {
    expect(relativeTime(null, NOW)).toBe("never");
    expect(relativeTime("last Tuesday", NOW)).toBe("unknown");
  });
});

describe("RelativeTime", () => {
  it("is a time element with the exact moment as its title", () => {
    const html = renderToStaticMarkup(
      createElement(RelativeTime, { value: "2026-10-24T11:55:00.000Z", now: NOW }),
    );
    expect(html).toMatch(
      /^<time dateTime="2026-10-24T11:55:00.000Z" title="[^"]+">5 min ago<\/time>$/,
    );
  });

  it("is plain muted text, not a time element, when there is nothing to date", () => {
    const html = renderToStaticMarkup(createElement(RelativeTime, { value: null, now: NOW }));
    expect(html).toBe('<span class="person-muted">never</span>');
  });
});

describe("SeverityDot", () => {
  const render = (severity: Parameters<typeof SeverityDot>[0]["severity"], count?: number) =>
    renderToStaticMarkup(createElement(SeverityDot, { severity, count }));

  it("names the severity next to its dot, so colour is never the only signal", () => {
    expect(render("critical")).toBe(
      '<span class="sev sev--critical"><span class="sev__dot" aria-hidden="true"></span>Critical</span>',
    );
    expect(render("high")).toContain("sev--high");
    expect(render("medium")).toContain(">Medium<");
  });

  it("reads as a count when it has one, and gives info the colour of low", () => {
    expect(render("critical", 2)).toContain(">2 critical<");
    expect(render("medium", 0)).toContain(">0 medium<");
    expect(render("info")).toContain("sev--low");
    expect(render("info")).toContain(">Info<");
  });
});

describe("SkeletonRows", () => {
  it("draws the rows hidden from assistive technology, with a label for it", () => {
    const html = renderToStaticMarkup(
      createElement(SkeletonRows, { rows: 3, label: "Loading audits" }),
    );
    expect(html.match(/skeleton__row/g)).toHaveLength(3);
    expect(html.match(/aria-hidden="true"/g)).toHaveLength(3);
    expect(html).toContain('<span class="visually-hidden">Loading audits</span>');
    expect(html).not.toContain("Loading…");
  });
});

const finding: Finding = {
  id: "66666666-6666-4666-8666-666666666666",
  module: "signins.audit",
  ruleId: "signins.unfamiliar_signin",
  ruleVersion: 1,
  purpose: "security",
  severity: "high",
  score: 80,
  subject: { kind: "person", id: "44444444-4444-4444-8444-444444444444", label: "Rowan Ellis" },
  title: "A sign-in from a new place",
  summary: "Rowan Ellis signed in from a place not seen before.",
  evidence: [
    {
      source: "Microsoft Entra sign-in log",
      ref: "e-1",
      observedAt: "2026-10-22T09:05:00.000Z",
      fields: { country: "Norway" },
    },
  ],
  limitations: ["Synthetic."],
  windowStart: "2026-10-22T09:05:00.000Z",
  windowEnd: "2026-10-22T09:05:00.000Z",
  state: "open",
  occurrences: 1234,
  firstSeen: "2026-10-20T09:06:00.000Z",
  lastSeen: "2026-10-22T09:06:00.000Z",
  stateChangedAt: null,
};

describe("FindingItem", () => {
  const draw = (value: Finding = finding) =>
    at("2026-10-22T12:00:00.000Z", () =>
      renderToStaticMarkup(
        createElement(FindingItem, {
          finding: value,
          busy: false,
          onState: () => {},
          onOpenPerson: () => {},
        }),
      ),
    );

  it("marks severity with the shared dot and name, on every page that lists findings", () => {
    const html = draw();
    expect(html).toContain(severityDot("high", "High"));
    expect(draw({ ...finding, severity: "critical" })).toContain(
      severityDot("critical", "Critical"),
    );
    expect(html).not.toContain("finding-severity");
  });

  it("says when it was last and first seen in words, in separate facts with no punctuation joining them", () => {
    const meta = draw().match(/<p class="finding__meta meta-line">(.*?)<\/p>/)?.[1] ?? "";
    expect(meta).toMatch(/^<span>Last seen /);
    expect(meta).toMatch(relativeCell("2026-10-22T09:06:00.000Z", "2 h ago"));
    expect(meta).toMatch(relativeCell("2026-10-20T09:06:00.000Z", "2 days ago"));
    expect(meta).toContain("<span>First seen ");
    expect(meta).toContain(`<span>${(1234).toLocaleString()} occurrences</span>`);
    expect(meta.match(/<span>/g)).toHaveLength(3);
    expect(meta).not.toMatch(/[\u00b7\u2014]/);
  });

  it("keeps the evidence inside a finding at its exact moment", () => {
    const evidence = draw().match(/<dl class="finding-evidence">(.*?)<\/dl>/)?.[1] ?? "";
    expect(evidence).toContain("<dt>Observed</dt>");
    expect(evidence).not.toContain("<time");
  });
});

// The stylesheet is plain CSS, so the rules the contract states are checked as text: they are
// easy to lose in an edit and nothing else would notice.
const css = readFileSync(new URL("../src/styles.css", import.meta.url), "utf8");

function rule(selector: string): string {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = css.match(new RegExp(`(?:^|\\n)${escaped}\\s*\\{([^}]*)\\}`));
  if (!match) throw new Error(`No rule for ${selector}`);
  return match[1] ?? "";
}

describe("stylesheet", () => {
  it("maps the four severities onto the existing status colours and adds no hue", () => {
    expect(rule(":root")).toContain("--sev-critical: var(--danger);");
    expect(rule(":root")).toContain("--sev-high: var(--danger);");
    expect(rule(":root")).toContain("--sev-medium: var(--warning);");
    expect(rule(":root")).toContain("--sev-low: var(--info);");
  });

  // Every table cell and metric value on every page gets its figures from this one rule. A page
  // that wants them adds nothing, and a second declaration would be a per-page copy.
  it("sets tabular figures once, on the root, and nowhere else", () => {
    expect(rule(":root")).toContain("font-variant-numeric: tabular-nums;");
    expect(css.match(/font-variant-numeric/g)).toHaveLength(1);
  });

  it("separates short facts with one shared hairline, the person header's included", () => {
    expect(rule(".meta-line")).toContain("display: flex;");
    const divider = parseRules(css).find((item) => item.selectors.includes(".meta-line > * + *"));
    expect(divider?.selectors).toContain(".person-header__role > span + span");
    expect(divider?.body).toContain("border-left: 1px solid var(--border-strong);");
  });

  it("keeps loading bars still", () => {
    for (const selector of [".skeleton", ".skeleton__row"]) {
      expect(rule(selector), selector).not.toMatch(/animation|transition/);
    }
    expect(css).not.toMatch(/\.skeleton[^{]*\{[^}]*animation/);
  });
});

// A small reader of the stylesheet's rules, enough to compare how specific two selectors are and
// to tell which rules sit inside a media query. It is not a cascade engine; the tests below name
// exactly what they compare.
interface CssRule {
  selectors: string[];
  body: string;
  media: string | null;
}

function parseRules(source: string, media: string | null = null): CssRule[] {
  const text = source.replace(/\/\*[\s\S]*?\*\//g, "");
  const rules: CssRule[] = [];
  let position = 0;
  while (position < text.length) {
    const open = text.indexOf("{", position);
    if (open === -1) break;
    const prelude = text.slice(position, open).trim();
    let depth = 1;
    let close = open + 1;
    while (close < text.length && depth > 0) {
      if (text[close] === "{") depth += 1;
      else if (text[close] === "}") depth -= 1;
      close += 1;
    }
    const inner = text.slice(open + 1, close - 1);
    if (prelude.startsWith("@media")) rules.push(...parseRules(inner, prelude.slice(6).trim()));
    else if (!prelude.startsWith("@"))
      rules.push({
        selectors: prelude.split(/,(?![^(]*\))/).map((selector) => selector.trim()),
        body: inner,
        media,
      });
    position = close;
  }
  return rules;
}

// [ids, classes and attributes and pseudo-classes, types and pseudo-elements]. :where() adds none.
function specificity(selector: string): [number, number, number] {
  let rest = selector.replace(/:where\([^)]*\)/g, "");
  const ids = rest.match(/#[\w-]+/g)?.length ?? 0;
  rest = rest.replace(/#[\w-]+/g, "");
  const attributes = rest.match(/\[[^\]]*\]/g)?.length ?? 0;
  rest = rest.replace(/\[[^\]]*\]/g, "");
  const pseudoElements = rest.match(/::[\w-]+/g)?.length ?? 0;
  rest = rest.replace(/::[\w-]+/g, "");
  const pseudoClasses = rest.match(/:[\w-]+(?:\([^)]*\))?/g)?.length ?? 0;
  rest = rest.replace(/:[\w-]+(?:\([^)]*\))?/g, "");
  const classes = rest.match(/\.[\w-]+/g)?.length ?? 0;
  rest = rest.replace(/\.[\w-]+/g, "");
  const types = rest.match(/[a-zA-Z][\w-]*/g)?.length ?? 0;
  return [ids, classes + attributes + pseudoClasses, types + pseudoElements];
}

const atLeast = (a: number[], b: number[]) => {
  for (let index = 0; index < a.length; index += 1) {
    if ((a[index] ?? 0) !== (b[index] ?? 0)) return (a[index] ?? 0) > (b[index] ?? 0);
  }
  return true;
};

describe("audits table on a phone", () => {
  const rules = parseRules(css);
  const phone = rules.filter((item) => item.media === "(max-width: 720px)");

  // The first stacked row once looked right and every later one had tall gaps: the rule that
  // removed the cells' padding matched only the first row of a group, and the one meant for the
  // rest was weaker than the desktop padding on th and td. A row is now one set of rules for all.
  it("removes the desktop padding from every row's cells with a rule as specific as the padding", () => {
    const reset = phone.find(
      (item) =>
        item.selectors.includes(".catalog-table .catalog-table__row > *") &&
        /padding:\s*0;/.test(item.body),
    );
    expect(reset).toBeDefined();
    const desktopPadding = rules
      .filter((item) => item.media === null && /(^|[\s;])padding(-[a-z]+)?:/.test(item.body))
      .flatMap((item) => item.selectors)
      .filter(
        (selector) =>
          selector.startsWith(".catalog-table") && !/thead|__group|skeleton/.test(selector),
      );
    // The two rules that pad the cells: th and td, and the name cell.
    expect(desktopPadding.length).toBeGreaterThanOrEqual(3);
    for (const selector of desktopPadding)
      expect(
        atLeast(specificity(".catalog-table .catalog-table__row > *"), specificity(selector)),
        selector,
      ).toBe(true);
  });

  it("uses no sibling rule to pad or space some rows and not others", () => {
    const uneven = phone.filter(
      (item) =>
        /padding|margin|gap/.test(item.body) &&
        item.selectors.some((selector) => selector.includes("+")),
    );
    expect(uneven).toEqual([]);
  });

  it("stacks each row as labelled lines with a small gap", () => {
    const row = phone.find((item) => item.selectors.includes(".catalog-table .catalog-table__row"));
    expect(row?.body).toContain("gap: 2px 12px;");
    expect(row?.body).toContain("padding: 10px 16px;");
    const label = phone.find((item) => item.selectors.includes(".catalog-table td[data-label]"));
    expect(label?.body).toContain("display: flex;");
    expect(label?.body).toContain("gap: 0 8px;");
  });

  it("measures specificity the way the cascade does", () => {
    expect(specificity(".a .b > *")).toEqual([0, 2, 0]);
    expect(specificity(".catalog-table th")).toEqual([0, 1, 1]);
    expect(specificity('.catalog-table tbody th[scope="row"]')).toEqual([0, 2, 2]);
    expect(specificity(".x:where(.y, .z) > li")).toEqual([0, 1, 1]);
    expect(specificity(".catalog-table td[data-label]::before")).toEqual([0, 2, 2]);
  });
});

function lab(color: string): [number, number, number] {
  const channel = (index: number) => {
    const value = Number.parseInt(color.slice(1 + index * 2, 3 + index * 2), 16) / 255;
    return value > 0.04045 ? ((value + 0.055) / 1.055) ** 2.4 : value / 12.92;
  };
  const [r, g, b] = [channel(0), channel(1), channel(2)];
  const xyz = [
    (0.4124 * r + 0.3576 * g + 0.1805 * b) / 0.95047,
    0.2126 * r + 0.7152 * g + 0.0722 * b,
    (0.0193 * r + 0.1192 * g + 0.9505 * b) / 1.08883,
  ].map((value) => (value > 0.008856 ? value ** (1 / 3) : 7.787 * value + 16 / 116));
  const [x = 0, y = 0, z = 0] = xyz;
  return [116 * y - 16, 500 * (x - y), 200 * (y - z)];
}

const colorOf = (token: string) => {
  const match = css.match(new RegExp(`${token}:\\s*(#[0-9a-f]{6})`));
  if (!match?.[1]) throw new Error(`No colour for ${token}`);
  return match[1];
};

describe("audit result colours", () => {
  it("separate critical from warning by more than a hue, side by side", () => {
    const [a, b] = [lab(colorOf("--danger")), lab(colorOf("--warning"))];
    // CIE76 distance; about 10 is plain to see, and the two tones are 30 apart, though almost equal
    // in lightness, so hue is the only colour cue. Critical therefore also goes heavier and wears a
    // ring on its dot.
    expect(Math.hypot(a[0] - b[0], a[1] - b[1], a[2] - b[2])).toBeGreaterThan(20);
    expect(rule(".audit-result__issue--critical")).toContain("--issue-color: var(--sev-critical);");
    expect(rule(".audit-result__issue--warning")).toContain("--issue-color: var(--sev-medium);");
    expect(rule(".audit-result__issue--critical")).toContain("font-weight: 600;");
    expect(css).toMatch(/\.sev__dot--critical \{\n  box-shadow: 0 0 0 3px var\(--danger-soft\);/);
    expect(rule(".sev__dot--medium")).not.toContain("box-shadow");
  });

  it("colours only the issue, and mutes the issue of a neutral or missing tone", () => {
    expect(rule(".audit-result")).toContain("color: var(--text);");
    expect(rule(".audit-result__issue")).toContain("color: var(--issue-color, var(--text-3));");
    expect(css).not.toMatch(/\.audit-result--(critical|warning|clear)\s*\{/);
  });
});

describe("Overview rows", () => {
  it("wrap the title to two lines and cut nothing before that", () => {
    const head = rule(".overview-row__head");
    expect(head).toContain("-webkit-line-clamp: 2;");
    expect(head).toContain("line-clamp: 2;");
    expect(head).not.toMatch(/nowrap|ellipsis/);
    expect(css).not.toMatch(/\.overview-row__title\s*\{[^}]*(nowrap|ellipsis)/);
  });

  it("keep the subject to one muted line beside the time on the right", () => {
    expect(rule(".overview-row__subject")).toContain("text-overflow: ellipsis;");
    expect(rule(".overview-row__meta")).toContain("justify-content: space-between;");
    expect(rule(".overview-row__meta")).toContain("color: var(--text-2);");
  });
});
