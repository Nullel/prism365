import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, relative } from "node:path";
import { describe, expect, it } from "vitest";
import { pageCopy } from "../src/App.tsx";
import { localDemoDashboard } from "../src/demo.ts";

// Rules that hold across every page, checked on the source and the stylesheet themselves. The
// per-page tests show what each page draws; these fail when a page drifts back to its own way,
// including a page that is added later.

const webSource = new URL("../src/", import.meta.url).pathname;
// The demo dashboard's synthetic activity text is written here and in the API's demo data; the
// web page shows both.
const demoData = new URL("../../api/src/demo-data.ts", import.meta.url).pathname;

function walk(directory: string): string[] {
  return readdirSync(directory).flatMap((name) => {
    const path = join(directory, name);
    return statSync(path).isDirectory() ? walk(path) : [path];
  });
}

const sources = walk(webSource).filter((path) => /\.(tsx?|css)$/.test(path));
const name = (path: string) => relative(webSource, path);

// Removes comments and leaves strings alone, so a rule can be about what the page says and not
// about what a comment explains. A quote only opens a string when it does not follow a letter,
// which keeps the apostrophe in "Microsoft's" from opening one in JSX text.
function withoutComments(source: string): string {
  let out = "";
  let index = 0;
  let quote: string | null = null;
  while (index < source.length) {
    const char = source[index] ?? "";
    const next = source[index + 1] ?? "";
    if (quote) {
      out += char;
      if (char === "\\") {
        out += next;
        index += 2;
        continue;
      }
      if (char === quote) quote = null;
      index += 1;
      continue;
    }
    if (char === "/" && next === "/" && source[index - 1] !== ":") {
      while (index < source.length && source[index] !== "\n") index += 1;
      continue;
    }
    if (char === "/" && next === "*") {
      const end = source.indexOf("*/", index + 2);
      index = end === -1 ? source.length : end + 2;
      continue;
    }
    if ((char === '"' || char === "'" || char === "`") && !/[\w$]/.test(source[index - 1] ?? "")) {
      quote = char;
    }
    out += char;
    index += 1;
  }
  return out;
}

const code = new Map(sources.map((path) => [path, withoutComments(readFileSync(path, "utf8"))]));
const scripts = [...code].filter(([path]) => /\.tsx?$/.test(path));

describe("the comment reader the rules below stand on", () => {
  it("drops comments and keeps strings, URLs and apostrophes in text", () => {
    const source = [
      "const a = 'x'; // gone",
      '/* gone too */ const b = "https://example.invalid/a//b";',
      "<p>Microsoft's directory</p> // gone",
      "{/* gone */}",
    ].join("\n");
    const result = withoutComments(source);
    expect(result).toContain("const a = 'x';");
    expect(result).toContain('"https://example.invalid/a//b"');
    expect(result).toContain("<p>Microsoft's directory</p>");
    expect(result).not.toContain("gone");
  });
});

describe("one visual language across the pages", () => {
  it("never joins visible text with a middle dot", () => {
    // The trend arrows on the demo dashboard are glyphs, not separators, and are not looked at.
    // Any middle dot, however it is written, is a join; the pages put short facts in their own
    // spans with a rule between them (the meta-line class).
    const dot = /·|\\u00b7|&middot;|&#0*183;|&#x0*b7;/i;
    const found = [...code]
      .filter(([, text]) => dot.test(text))
      .map(([path]) => name(path))
      .concat(readFileSync(demoData, "utf8").match(dot) ? ["api/src/demo-data.ts"] : []);
    expect(found).toEqual([]);
  });

  it("carries no middle dot in the bundled demo dashboard's text either", () => {
    for (const activity of localDemoDashboard.recentActivity)
      expect(activity.evidence, activity.title).not.toMatch(/·/);
  });

  it("shows no page's own badge for severity: every severity is the shared dot and name", () => {
    const retired = /finding-severity|severity severity--/;
    expect(scripts.filter(([, text]) => retired.test(text)).map(([path]) => name(path))).toEqual(
      [],
    );
    const styles = code.get(join(webSource, "styles.css")) ?? "";
    expect(styles).not.toMatch(/\.finding-severity|\.severity\b/);
  });

  it("colours severity only with the shared tokens, never a colour of its own", () => {
    const styles = code.get(join(webSource, "styles.css")) ?? "";
    const rules = [...styles.matchAll(/([^{}]+)\{([^{}]*)\}/g)].map((match) => ({
      selector: (match[1] ?? "").trim(),
      body: match[2] ?? "",
    }));
    const severityRules = rules.filter(({ selector }) =>
      /\.sev\b|\.sev[-_]|marker--(critical|high|medium|low|info)|tile--(critical|warning)/.test(
        selector,
      ),
    );
    expect(severityRules.length).toBeGreaterThan(8);
    for (const { selector, body } of severityRules)
      expect(body, selector).not.toMatch(/#[0-9a-f]{3,8}\b|rgba?\(|hsla?\(/i);
  });

  it("has no loading text in a page: loading is skeleton rows", () => {
    // JSX text that starts with Loading, wherever it sits in the markup.
    const shown = scripts
      .filter(([, text]) => />\s*Loading[^<{]*</.test(text))
      .map(([path]) => name(path));
    expect(shown).toEqual([]);
    // The ellipsis form as a string is for a button that is busy ("Show older changes" while it
    // loads them) and for the quick search's own notes. Nothing else may use it.
    const allowed: Record<string, number> = {
      "Changes.tsx": 1,
      "Handoffs.tsx": 1,
      "PersonChanges.tsx": 1,
      "QuickSearch.tsx": 2,
    };
    const counts = Object.fromEntries(
      scripts
        .map(([path, text]) => [name(path), text.match(/Loading[^"`<]*…/g)?.length ?? 0] as const)
        .filter(([, count]) => count > 0),
    );
    expect(counts).toEqual(allowed);
  });

  it("has one empty-state style: the empty-data line", () => {
    const retired = /signin-findings__empty|audit-empty-run/;
    expect([...code].filter(([, text]) => retired.test(text)).map(([path]) => name(path))).toEqual(
      [],
    );
  });

  it("writes every page description as one short sentence", () => {
    const entries = Object.entries(pageCopy);
    expect(entries.length).toBeGreaterThanOrEqual(10);
    for (const [view, { description }] of entries) {
      expect(description, view).toMatch(/^[A-Z][^.:;!?—]*\.$/);
      expect(description.length, view).toBeLessThanOrEqual(90);
      expect(description.split(/\s+/).length, view).toBeLessThanOrEqual(14);
    }
  });
});
