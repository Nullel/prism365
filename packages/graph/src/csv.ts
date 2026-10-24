// A strict, bounded RFC 4180 reader for Microsoft's usage report downloads (ADR-0010), which are
// untrusted provider text. It scans the input one character at a time and no regular expression
// ever sees it, so the time it takes is linear in the input, and every count is capped: columns,
// rows and field length. Values come back exactly as written, with no type guessing; callers find
// a column through `index` by header name, never by a fixed position, because Microsoft adds and
// reorders report columns.

export const CSV_MAX_COLUMNS = 100;
export const CSV_MAX_ROWS = 50_000;
export const CSV_MAX_FIELD_LENGTH = 8_192;

export interface CsvTable {
  header: string[];
  // Each header name's position in a row. A column the report does not have is not in it, so it
  // reads as undefined rather than as a position.
  index: ReadonlyMap<string, number>;
  // One array of values per kept data row, as long as the header. Rows are arrays, not one object
  // per row, because objects keyed by 100 names cost many times the report's own size.
  rows: string[][];
  // Data rows left out: a field count that differs from the header's, a field over the length
  // limit, or broken quoting (a quote inside an unquoted field, text after a closing quote, or a
  // carriage return without a line feed).
  skippedRows: number;
  // true when the input has more data rows than the row limit; the rest was not read.
  truncated: boolean;
}

// Fixed messages: none of them repeats any of the input.
const problems = {
  no_header: "the report has no header row",
  header_quoting: "the header row has broken quoting",
  header_length: "a header name is longer than the field limit",
  columns: `the report has more than ${CSV_MAX_COLUMNS} columns`,
  empty_header: "a header name is empty",
  duplicate_header: "two columns have the same header name",
  unterminated_quote: "a quoted field is not closed before the end of the report",
} as const;

export type CsvProblem = keyof typeof problems;

export class CsvFormatError extends Error {
  readonly problem: CsvProblem;
  constructor(problem: CsvProblem) {
    super(`The CSV could not be read: ${problems[problem]}.`);
    this.name = "CsvFormatError";
    this.problem = problem;
  }
}

const QUOTE = 0x22;
const COMMA = 0x2c;
const CR = 0x0d;
const LF = 0x0a;
const BOM = 0xfeff;

type Flaw = "quoting" | "length" | "fields";

interface ScannedRecord {
  fields: string[];
  flaw: Flaw | null;
  // Index just past the record's line ending.
  next: number;
}

// Whether a field ends at `at`: a comma, a line feed, a carriage return and line feed, a carriage
// return as the last character, or the end of the text.
function endsField(text: string, at: number): boolean {
  if (at >= text.length) return true;
  const code = text.charCodeAt(at);
  if (code === COMMA || code === LF) return true;
  return code === CR && (at + 1 >= text.length || text.charCodeAt(at + 1) === LF);
}

// Index just past the empty lines starting at `at`. An empty line holds no field at all, so it is
// not a record; an empty value is written as "".
function afterEmptyLines(text: string, at: number): number {
  let position = at;
  for (;;) {
    const code = text.charCodeAt(position);
    if (code === LF) position += 1;
    else if (code === CR && text.charCodeAt(position + 1) === LF) position += 2;
    else if (code === CR && position + 1 === text.length) position += 1;
    else return position;
  }
}

// Reads the record that starts at `start`, keeping at most `maxFields` fields. A flaw does not stop
// the scan: the record is read to its end so the next one starts in the right place, and the
// caller decides what the flaw means.
function scanRecord(text: string, start: number, maxFields: number): ScannedRecord {
  const fields: string[] = [];
  let flaw: Flaw | null = null;
  let at = start;
  for (;;) {
    let value = "";
    let length = 0;
    if (text.charCodeAt(at) === QUOTE) {
      // A quoted field runs to the next quote that is not doubled; commas and line breaks inside
      // it belong to the value, and a doubled quote is one quote.
      at += 1;
      let from = at;
      for (;;) {
        if (at >= text.length) throw new CsvFormatError("unterminated_quote");
        if (text.charCodeAt(at) !== QUOTE) {
          at += 1;
          continue;
        }
        length += at - from;
        if (length <= CSV_MAX_FIELD_LENGTH) value += text.slice(from, at);
        if (text.charCodeAt(at + 1) === QUOTE) {
          length += 1;
          if (length <= CSV_MAX_FIELD_LENGTH) value += '"';
          at += 2;
          from = at;
          continue;
        }
        at += 1;
        break;
      }
      // Anything between the closing quote and the end of the field is broken quoting. It is
      // read as plain text, so a quote in it does not open a new quoted section.
      if (!endsField(text, at)) {
        flaw ??= "quoting";
        while (!endsField(text, at)) at += 1;
      }
    } else {
      const from = at;
      while (!endsField(text, at)) {
        // A quote inside an unquoted field, or a carriage return that does not end the line.
        const code = text.charCodeAt(at);
        if (code === QUOTE || code === CR) flaw ??= "quoting";
        at += 1;
      }
      length = at - from;
      if (length <= CSV_MAX_FIELD_LENGTH) value = text.slice(from, at);
    }
    if (length > CSV_MAX_FIELD_LENGTH) flaw ??= "length";
    if (fields.length < maxFields) fields.push(value);
    else flaw ??= "fields";
    if (at >= text.length) return { fields, flaw, next: at };
    const code = text.charCodeAt(at);
    if (code === COMMA) {
      at += 1;
      continue;
    }
    if (code === CR) at += 1;
    if (text.charCodeAt(at) === LF) at += 1;
    return { fields, flaw, next: at };
  }
}

function readHeader(text: string, start: number): { header: string[]; next: number } {
  if (start >= text.length) throw new CsvFormatError("no_header");
  // Fields past the column limit are not kept; the first one is enough to refuse the header.
  const scanned = scanRecord(text, start, CSV_MAX_COLUMNS);
  if (scanned.flaw === "fields") throw new CsvFormatError("columns");
  if (scanned.flaw === "length") throw new CsvFormatError("header_length");
  if (scanned.flaw === "quoting") throw new CsvFormatError("header_quoting");
  const seen = new Set<string>();
  for (const name of scanned.fields) {
    if (name.trim() === "") throw new CsvFormatError("empty_header");
    if (seen.has(name)) throw new CsvFormatError("duplicate_header");
    seen.add(name);
  }
  return { header: scanned.fields, next: scanned.next };
}

// Parses a whole CSV document: one leading byte-order mark is dropped, lines end in CRLF or LF,
// and the first record is the header. A data row that cannot be read is skipped and counted; only
// a problem that leaves the columns or the row boundaries unknown (the header, or a quote that is
// never closed) fails the whole document.
export function parseCsv(input: string): CsvTable {
  const text = input.charCodeAt(0) === BOM ? input.slice(1) : input;
  const { header, next } = readHeader(text, afterEmptyLines(text, 0));
  const rows: string[][] = [];
  let skippedRows = 0;
  let truncated = false;
  let at = next;
  for (;;) {
    at = afterEmptyLines(text, at);
    if (at >= text.length) break;
    if (rows.length >= CSV_MAX_ROWS) {
      truncated = true;
      break;
    }
    const scanned = scanRecord(text, at, header.length);
    at = scanned.next;
    if (scanned.flaw !== null || scanned.fields.length !== header.length) {
      skippedRows += 1;
      continue;
    }
    rows.push(scanned.fields);
  }
  return {
    header,
    index: new Map(header.map((name, position) => [name, position])),
    rows,
    skippedRows,
    truncated,
  };
}
