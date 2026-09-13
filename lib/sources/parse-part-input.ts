import type { ExtraFields } from "@/lib/types";
import type { PartFields } from "./sharepoint-excel-source";

/**
 * Fields a technician submits from the "Add a part" / "Edit part" forms
 * (see components/AddPartForm.tsx and components/PartDetail.tsx),
 * validated before either updatePartInWorkbook or addPartToWorkbook
 * ever sees them. `sourceId` says which SharePoint workbook to write
 * to; `existingPartNumber` is present only for an update (it's how the
 * row to overwrite gets found — see parsePartUpdateInput) and absent
 * for adding a brand-new part.
 */
export interface ParsedPartInput {
  sourceId: string;
  fields: PartFields;
}

export interface ParsedPartUpdateInput extends ParsedPartInput {
  existingPartNumber: string;
  /** Which row to overwrite among every row sharing `existingPartNumber`
   *  — see RawPart.partNumberOccurrence and updatePartInWorkbook. Defaults
   *  to 1 (the common case: only one row has this part_number) when the
   *  client doesn't send one. */
  existingPartOccurrence: number;
}

/**
 * Validates and normalizes a new-part submission. `part_number` and
 * `sourceId` are required and non-blank; `quantity_on_hand` must be a
 * finite number and is clamped to 0 if negative (a technician fat-
 * fingering "-5" should not write a negative stock count); description
 * and bin_location default to "" when omitted, matching how a source
 * with no such column configured reads back (see mapTableRowsToRawParts).
 */
export function parsePartInput(body: unknown): { input: ParsedPartInput } | { error: string } {
  if (!body || typeof body !== "object") {
    return { error: "Invalid request body." };
  }
  const record = body as Record<string, unknown>;

  if (typeof record.sourceId !== "string" || !record.sourceId.trim()) {
    return { error: "Missing or empty field: sourceId" };
  }
  if (typeof record.part_number !== "string" || !record.part_number.trim()) {
    return { error: "Missing or empty field: part_number" };
  }
  const quantity = Number(record.quantity_on_hand);
  if (!Number.isFinite(quantity)) {
    return { error: "quantity_on_hand must be a number." };
  }

  return {
    input: {
      sourceId: record.sourceId.trim(),
      fields: {
        part_number: record.part_number.trim(),
        description: normalizeOptionalText(record.description),
        bin_location: normalizeOptionalText(record.bin_location),
        quantity_on_hand: Math.max(0, quantity),
        category: normalizeOptionalText(record.category),
        extraFields: normalizeExtraFields(record.extraFields),
      },
    },
  };
}

/** Passes an "Edit this part" submission's `extraFields` through to
 *  PartFields, re-checking each entry's shape rather than trusting the
 *  client's JSON wholesale — a malformed or missing `kind`/`value` on
 *  one entry falls back to a safe default instead of writing `undefined`
 *  or some unexpected type into the workbook. See ExtraField in
 *  lib/types.ts for what each kind means; `options` (only meaningful
 *  for "select") is preserved when present so the value keeps whatever
 *  fixed choice list it was classified with. */
function normalizeExtraFields(value: unknown): ExtraFields {
  if (!value || typeof value !== "object") {
    return {};
  }
  const result: ExtraFields = {};
  for (const [header, raw] of Object.entries(value as Record<string, unknown>)) {
    if (!raw || typeof raw !== "object" || typeof (raw as Record<string, unknown>).kind !== "string") {
      continue;
    }
    const entry = raw as Record<string, unknown>;
    const options = Array.isArray(entry.options)
      ? entry.options.filter((option): option is string => typeof option === "string")
      : undefined;
    if (entry.kind === "boolean") {
      result[header] = { kind: "boolean", value: Boolean(entry.value) };
    } else if (entry.kind === "date") {
      result[header] = { kind: "date", value: typeof entry.value === "string" ? entry.value : "" };
    } else if (entry.kind === "select") {
      result[header] = {
        kind: "select",
        value: typeof entry.value === "string" ? entry.value : "",
        options,
      };
    } else {
      result[header] = {
        kind: "text",
        value: typeof entry.value === "string" ? entry.value : String(entry.value ?? ""),
      };
    }
  }
  return result;
}

/**
 * Same validation as parsePartInput, plus a required
 * `existingPartNumber` — the part number to locate and overwrite,
 * which may differ from the new `part_number` when a technician is
 * correcting a typo'd part number as part of the same edit.
 */
export function parsePartUpdateInput(
  body: unknown,
): { input: ParsedPartUpdateInput } | { error: string } {
  const parsed = parsePartInput(body);
  if ("error" in parsed) {
    return parsed;
  }

  const record = body as Record<string, unknown>;
  if (typeof record.existingPartNumber !== "string" || !record.existingPartNumber.trim()) {
    return { error: "Missing or empty field: existingPartNumber" };
  }

  const occurrence = Number(record.existingPartOccurrence);
  const existingPartOccurrence =
    Number.isFinite(occurrence) && occurrence >= 1 ? Math.trunc(occurrence) : 1;

  return {
    input: {
      ...parsed.input,
      existingPartNumber: record.existingPartNumber.trim(),
      existingPartOccurrence,
    },
  };
}

function normalizeOptionalText(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}
