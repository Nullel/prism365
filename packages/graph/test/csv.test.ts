import { describe, expect, it } from "vitest";
import {
  CSV_MAX_COLUMNS,
  CSV_MAX_FIELD_LENGTH,
  CSV_MAX_ROWS,
  CsvFormatError,
  parseCsv,
} from "../src/csv.ts";

// Every value here is invented: .example addresses and made-up people.

const problem = (input: string) => {
  try {
    parseCsv(input);
  } catch (error) {
    if (error instanceof CsvFormatError) return error.problem;
    throw error;
  }
  return null;
};

describe("parseCsv", () => {
  it("strips one leading byte-order mark and keeps a second as text", () => {
    expect(parseCsv("\uFEFFName,Count\r\nAda,1\r\n").header).toEqual(["Name", "Count"]);
    expect(parseCsv("Name,Count\r\nAda,1\r\n").header).toEqual(["Name", "Count"]);
    expect(parseCsv("\uFEFF\uFEFFName\n\uFEFFAda\n")).toMatchObject({
      header: ["\uFEFFName"],
      rows: [{ "\uFEFFName": "\uFEFFAda" }],
    });
  });

  it("reads CRLF, LF and mixed line endings, with or without a final one", () => {
    const expected = [
      { Name: "Ada", Count: "1" },
      { Name: "Brook", Count: "2" },
    ];
    for (const input of [
      "Name,Count\r\nAda,1\r\nBrook,2\r\n",
      "Name,Count\nAda,1\nBrook,2\n",
      "Name,Count\r\nAda,1\nBrook,2",
      "Name,Count\nAda,1\r\nBrook,2\r",
    ])
      expect(parseCsv(input), JSON.stringify(input)).toEqual({
        header: ["Name", "Count"],
        rows: expected,
        skippedRows: 0,
        truncated: false,
      });
  });

  it("reads quoted commas, doubled quotes and line breaks inside quotes exactly as written", () => {
    const table = parseCsv(
      'Display Name,Note,Products\r\n"Okafor, Ada","She said ""hi""","Line one\r\nLine two\nthree"\r\n',
    );
    expect(table.skippedRows).toBe(0);
    expect(table.rows).toEqual([
      {
        "Display Name": "Okafor, Ada",
        Note: 'She said "hi"',
        Products: "Line one\r\nLine two\nthree",
      },
    ]);
    // A quoted header, a field that is only a quote, and a doubled quote at either end.
    expect(parseCsv('"A ""quoted"" name",B\n"""","""x"""\n').rows).toEqual([
      { 'A "quoted" name': '"', B: '"x"' },
    ]);
  });

  it("keeps values as strings with no trimming or type guessing", () => {
    const table = parseCsv(
      "Is Deleted,Count,Date,Padded\nFalse,0012, 2026-10-01 ,  spaced  \nTrue,1e3,,\n",
    );
    expect(table.rows).toEqual([
      { "Is Deleted": "False", Count: "0012", Date: " 2026-10-01 ", Padded: "  spaced  " },
      { "Is Deleted": "True", Count: "1e3", Date: "", Padded: "" },
    ]);
  });

  it("reads empty fields, quoted empty fields and a trailing empty field", () => {
    expect(parseCsv('A,B,C\n,,\n"",x,""\na,b,\n').rows).toEqual([
      { A: "", B: "", C: "" },
      { A: "", B: "x", C: "" },
      { A: "a", B: "b", C: "" },
    ]);
  });

  it("returns no rows for a header-only report", () => {
    for (const input of ["A,B", "A,B\r\n", "\uFEFFA,B\n\n\n"])
      expect(parseCsv(input)).toEqual({
        header: ["A", "B"],
        rows: [],
        skippedRows: 0,
        truncated: false,
      });
  });

  it("ignores empty lines, which hold no field, but keeps a row of one quoted empty value", () => {
    expect(parseCsv("A\n\nx\r\n\r\n\ny\n")).toMatchObject({
      rows: [{ A: "x" }, { A: "y" }],
      skippedRows: 0,
    });
    expect(parseCsv('A\n""\n').rows).toEqual([{ A: "" }]);
    expect(parseCsv("\n\nA,B\n1,2\n").header).toEqual(["A", "B"]);
  });

  it("requires a header row", () => {
    for (const input of ["", "\uFEFF", "\n", "\r\n\r\n"])
      expect(problem(input), JSON.stringify(input)).toBe("no_header");
  });

  it("refuses duplicate, empty and blank header names", () => {
    expect(problem("Name,Count,Name\n1,2,3\n")).toBe("duplicate_header");
    expect(problem('Name,"Name"\n')).toBe("duplicate_header");
    expect(problem("Name,,Count\n")).toBe("empty_header");
    expect(problem("Name,Count,\n")).toBe("empty_header");
    expect(problem('Name,""\n')).toBe("empty_header");
    expect(problem("Name,   \n")).toBe("empty_header");
    // Names differing only in case or spacing are different columns.
    expect(parseCsv("Name,name,Name \n1,2,3\n").header).toEqual(["Name", "name", "Name "]);
  });

  it("refuses a header with broken quoting or an overlong name", () => {
    expect(problem('Na"me,Count\n')).toBe("header_quoting");
    expect(problem('"Name"x,Count\n')).toBe("header_quoting");
    expect(problem("Name\rCount\n")).toBe("header_quoting");
    expect(problem(`${"N".repeat(CSV_MAX_FIELD_LENGTH + 1)},Count\n`)).toBe("header_length");
    expect(parseCsv(`${"N".repeat(CSV_MAX_FIELD_LENGTH)},Count\n`).header[0]).toHaveLength(
      CSV_MAX_FIELD_LENGTH,
    );
  });

  it("allows 100 columns and refuses 101", () => {
    const header = (count: number) =>
      Array.from({ length: count }, (_, index) => `Column ${index}`).join(",");
    const row = (count: number) => Array.from({ length: count }, () => "v").join(",");
    const widest = parseCsv(`${header(CSV_MAX_COLUMNS)}\n${row(CSV_MAX_COLUMNS)}\n`);
    expect(widest.header).toHaveLength(CSV_MAX_COLUMNS);
    expect(widest.rows).toHaveLength(1);
    expect(problem(`${header(CSV_MAX_COLUMNS + 1)}\n`)).toBe("columns");
  });

  it("skips and counts rows whose field count differs from the header", () => {
    const table = parseCsv("A,B,C\n1,2,3\n1,2\n1,2,3,4\nx\n4,5,6\n,,,\n");
    expect(table.rows).toEqual([
      { A: "1", B: "2", C: "3" },
      { A: "4", B: "5", C: "6" },
    ]);
    expect(table.skippedRows).toBe(4);
    // A ragged row with a quoted line break is still one row.
    expect(parseCsv('A,B\n"x\ny"\n1,2\n')).toMatchObject({
      rows: [{ A: "1", B: "2" }],
      skippedRows: 1,
    });
  });

  it("skips and counts rows with broken quoting, and finds the next row after them", () => {
    const table = parseCsv(
      'A,B\nab"c,1\n"ab"c,2\n"ab" ,3\nx\ry,4\n"ok, really",5\n"ab"c"d,e",6\nlast,7\n',
    );
    expect(table.rows).toEqual([
      { A: "ok, really", B: "5" },
      { A: "last", B: "7" },
    ]);
    expect(table.skippedRows).toBe(5);
  });

  it("skips a row with a field over 8,192 characters and keeps one at the limit", () => {
    const atLimit = "a".repeat(CSV_MAX_FIELD_LENGTH);
    const over = "b".repeat(CSV_MAX_FIELD_LENGTH + 1);
    // Quoted fields count their value, with each doubled quote as one character.
    const quotedAtLimit = `"${'"'.repeat(2 * 10)}${"c".repeat(CSV_MAX_FIELD_LENGTH - 10)}"`;
    const quotedOver = `"${"d".repeat(CSV_MAX_FIELD_LENGTH)}"""`;
    const table = parseCsv(
      `A,B\n${atLimit},1\n${over},2\n${quotedAtLimit},3\n${quotedOver},4\n"${over}",5\nend,6\n`,
    );
    expect(table.rows.map((row) => row.B)).toEqual(["1", "3", "6"]);
    expect(table.rows[0]?.A).toBe(atLimit);
    expect(table.rows[1]?.A).toBe(`${'"'.repeat(10)}${"c".repeat(CSV_MAX_FIELD_LENGTH - 10)}`);
    expect(table.skippedRows).toBe(3);
  });

  it("reads 50,000 rows and marks a report with more as truncated", () => {
    const rows = (count: number) =>
      `Id,Value\n${Array.from({ length: count }, (_, index) => `${index},v`).join("\n")}\n`;
    const full = parseCsv(rows(CSV_MAX_ROWS));
    expect(full.rows).toHaveLength(CSV_MAX_ROWS);
    expect(full.truncated).toBe(false);
    const more = parseCsv(rows(CSV_MAX_ROWS + 1));
    expect(more.rows).toHaveLength(CSV_MAX_ROWS);
    expect(more.rows.at(-1)?.Id).toBe(String(CSV_MAX_ROWS - 1));
    expect(more.truncated).toBe(true);
    // Skipped rows do not count toward the limit, and trailing empty lines are not more rows.
    const withSkips = parseCsv(`${rows(CSV_MAX_ROWS).replace("\n", "\nbad\n")}\n\n`);
    expect(withSkips).toMatchObject({ skippedRows: 1, truncated: false });
    expect(withSkips.rows).toHaveLength(CSV_MAX_ROWS);
  });

  it("fails on a quote that is never closed", () => {
    expect(problem('A,B\n1,"open\n2,3\n')).toBe("unterminated_quote");
    expect(problem('A,B\n1,"')).toBe("unterminated_quote");
    expect(problem('A,"B')).toBe("unterminated_quote");
    expect(problem('A,B\n1,"closed"""')).toBeNull();
  });

  it("gives errors fixed messages that repeat none of the input", () => {
    try {
      parseCsv("ada@lab.example,ada@lab.example\n");
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(CsvFormatError);
      expect(String(error)).not.toContain("ada@lab.example");
    }
  });

  it("keys records by header name without a prototype", () => {
    const table = parseCsv("__proto__,constructor,toString\n1,2,3\n");
    const [row] = table.rows;
    expect(Object.getPrototypeOf(row)).toBeNull();
    expect(Object.entries(row ?? {})).toEqual([
      ["__proto__", "1"],
      ["constructor", "2"],
      ["toString", "3"],
    ]);
    expect(row?.constructor).toBe("2");
    expect(parseCsv("A\n1\n").rows[0]?.toString).toBeUndefined();
  });

  it("reads a 10 MB report in well under a second", () => {
    const header = Array.from({ length: 20 }, (_, index) => `Column ${index}`).join(",");
    const line = Array.from({ length: 20 }, (_, index) =>
      index % 4 === 0 ? `"Okafor, Ada ""${index}"" of the lab"` : `synthetic-value-${index}`,
    ).join(",");
    const count = Math.ceil(10_000_000 / (line.length + 2));
    const input = `\uFEFF${header}\r\n${`${line}\r\n`.repeat(count)}`;
    expect(input.length).toBeGreaterThan(10_000_000);
    expect(count).toBeLessThan(CSV_MAX_ROWS);
    const started = performance.now();
    const table = parseCsv(input);
    const elapsed = performance.now() - started;
    expect(table.rows).toHaveLength(count);
    expect(table.rows[0]?.["Column 0"]).toBe('Okafor, Ada "0" of the lab');
    expect(elapsed).toBeLessThan(1_000);
  });

  it("stays linear on input built to make a backtracking reader slow", () => {
    const quotes = `A\n${'"'.repeat(2_000_001)}`;
    const started = performance.now();
    expect(problem(quotes)).toBe("unterminated_quote");
    const commas = parseCsv(`A,B\n${",".repeat(2_000_000)}\n`);
    expect(commas.skippedRows).toBe(1);
    expect(performance.now() - started).toBeLessThan(1_000);
  });
});
