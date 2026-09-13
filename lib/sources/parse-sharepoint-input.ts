import type { NewSharePointSource } from "@/lib/sources/sharepoint-source-store";

const REQUIRED_FIELDS: (keyof NewSharePointSource)[] = [
  "label",
  "siteHostname",
  "sitePath",
  "filePath",
  "tableName",
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
      siteHostname: (record.siteHostname as string).trim(),
      sitePath: (record.sitePath as string).trim(),
      filePath: (record.filePath as string).trim(),
      tableName: (record.tableName as string).trim(),
    },
  };
}
