// Provider text as stored: NUL removed (PostgreSQL text and jsonb refuse it) and lone UTF-16
// surrogates replaced, so one odd value from Microsoft can never make a record or a result
// impossible to save. Cut to length before cleaning, never after: cutting can split a surrogate
// pair, and cleaning then repairs the half that is left.
export function cleanText(value: string): string {
  return value
    .replaceAll("\u0000", "")
    .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, "�");
}
