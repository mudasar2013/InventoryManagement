import type { NewSharePointSource } from "@/lib/sources/sharepoint-source-store";

const REQUIRED_FIELDS: (keyof NewSharePointSource)[] = [
  "label",
  "siteHostname",
  "sitePath",
  "filePath",
  "tableName",
  "partNumberColumn",
  "quantityColumn",
];

/**
 * Validates and trims a SharePoint source submission from the "Data
 * sources" page. Shared by the collection route's POST (add) and the
 * [id] route's PATCH (edit) — a source's shape doesn't change between
 * creating and editing it, just which store function the parsed fields
 * get handed to afterwards.
 *
 * Deliberately NOT exported from either route.ts: Next.js route files
 * are only allowed to export recognized HTTP method handlers and a
 * short list of route-segment config options, so a helper like this
 * has to live in a plain module instead.
 */
export function parseSharePointSourceInput(
  body: unknown,
): { input: NewSharePointSource } | { error: string } {
  if (!body || typeof body !== "object") {
    return { error: "Invalid request body." };
  }

  const record = body as Record<string, unknown>;
  const missing = REQUIRED_FIELDS.filter(
    (field) => typeof record[field] !== "string" || !(record[field] as string).trim(),
  );
  if (missing.length > 0) {
    return { error: `Missing or empty fields: ${missing.join(", ")}` };
  }

  return {
    input: {
      label: (record.label as string).trim(),
      siteHostname: normalizeSiteHostname(record.siteHostname as string),
      sitePath: normalizeSitePath(record.sitePath as string),
      filePath: (record.filePath as string).trim(),
      tableName: (record.tableName as string).trim(),
      partNumberColumn: (record.partNumberColumn as string).trim(),
      quantityColumn: (record.quantityColumn as string).trim(),
      descriptionColumn: normalizeOptionalField(record.descriptionColumn),
      binLocationColumn: normalizeOptionalField(record.binLocationColumn),
      idColumn: normalizeOptionalField(record.idColumn),
    },
  };
}

/** Trims an optional column-name field, treating a blank string the
 *  same as it not being submitted at all (undefined) rather than
 *  storing empty strings — mapTableRowsToRawParts treats an undefined
 *  column name as "this field isn't in the sheet", which is exactly
 *  what leaving the field blank means. */
function normalizeOptionalField(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

/**
 * Strips a pasted-in "https://" (or "http://") scheme and any trailing
 * slash from the site hostname field. It's meant to hold a bare host
 * like "contoso.sharepoint.com" — the Graph call builds
 * `/sites/${siteHostname}:${sitePath}`, so a full URL pasted in by
 * mistake (very easy to do — it's literally what's in the browser's
 * address bar when you're looking at the site) produces a mangled path
 * like `/sites/https://contoso.sharepoint.com:/sites/Foo` that fails
 * before any request even goes out, surfacing only as a bare
 * "TypeError: fetch failed" with no useful cause. Silently fixing the
 * common mistake beats making someone puzzle that out from a generic
 * network error.
 */
function normalizeSiteHostname(value: string): string {
  return value.trim().replace(/^[a-z][a-z0-9+.-]*:\/\//i, "").replace(/\/+$/, "");
}

/** Site path should start with "/" and carry no trailing slash — trims
 *  whitespace and a trailing slash, and adds the leading "/" back if a
 *  pasted-in value (e.g. copied from a URL's path segment) omitted it. */
function normalizeSitePath(value: string): string {
  const trimmed = value.trim().replace(/\/+$/, "");
  if (!trimmed) {
    return trimmed;
  }
  return trimmed.startsWith("/") ? trimmed : `/${trimmed}`;
}
