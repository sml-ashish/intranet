/**
 * Small helpers for reading SharePoint list rows whose columns were created by hand
 * (their internal names are fixed at creation, e.g. "Sort order" becomes Sort_x0020_order).
 */

/** First non-empty value among the possible internal names. */
export function pickValue(row: Record<string, unknown>, names: string[]): unknown {
  for (const name of names) {
    const value: unknown = row[name];
    if (value) {
      return value;
    }
  }
  return undefined;
}

/** Text from a plain-text column, or the address from a Hyperlink column. */
export function asString(value: unknown): string {
  if (typeof value === 'string') {
    return value;
  }
  if (value && typeof value === 'object' && typeof (value as { Url?: unknown }).Url === 'string') {
    return (value as { Url: string }).Url;
  }
  return '';
}

/** A "Copy link" sharing address (.../:i:/r/sites/x/file.jpg?d=...) is a web page, not an image; turn it into the direct file address. */
export function toDirectImageUrl(url: string): string {
  const match: RegExpMatchArray | undefined = url.match(/^(https?:\/\/[^/]+)\/:[a-z]:\/[a-z](\/[^?]*)/i) || undefined;
  return match ? `${match[1]}${match[2]}` : url;
}