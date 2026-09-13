import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { describeError } from "@/lib/describe-error";
import { authOptions } from "@/lib/auth/options";
import { resolveSharePointConfigById } from "@/lib/sources/resolve-sharepoint-source";
import {
  COLUMN_MAP,
  connectToWorkbook,
  updatePartInWorkbook,
  type PartFields,
} from "@/lib/sources/sharepoint-excel-source";
import { getTagsForParts, setTagsForPart } from "@/lib/tags-store";

interface BulkFieldUpdate {
  /** The client's Part.id — carried through purely so the response can
   *  tell the client which specific selected part a result belongs to;
   *  never used to locate the row in the sheet (existingPartNumber +
   *  existingPartOccurrence do that). Needed because more than one
   *  selected part can share a part_number (see
   *  RawPart.partNumberOccurrence), so part_number alone can't
   *  distinguish which result is which back on the client. */
  id: string;
  sourceId: string;
  existingPartNumber: string;
  /** See ParsedPartUpdateInput.existingPartOccurrence — which row to
   *  overwrite when existingPartNumber matches more than one. */
  existingPartOccurrence: number;
  fields: PartFields;
}

interface BulkResult {
  id: string;
  part_number: string;
  ok: boolean;
  error?: string;
}

function isValidFields(value: unknown): value is PartFields {
  if (!value || typeof value !== "object") return false;
  const f = value as Record<string, unknown>;
  return (
    typeof f.part_number === "string" &&
    typeof f.description === "string" &&
    typeof f.bin_location === "string" &&
    typeof f.quantity_on_hand === "number" &&
    (f.category === undefined || typeof f.category === "string") &&
    (f.extraFields === undefined || typeof f.extraFields === "object")
  );
}

/**
 * Applies field changes to many parts at once (see
 * components/BulkEditPanel.tsx). The client sends each part's *full*
 * current field set per update — same shape as a normal single-part
 * PATCH /api/parts, just for several parts in one request — with only
 * the field(s) the technician actually chose to bulk-change overridden
 * on the client before sending. This route deliberately does NOT
 * accept partial/sparse fields: writing only a changed field while
 * leaving the rest undefined would blank out every field this request
 * doesn't mention (updatePartInWorkbook writes exactly the fields it's
 * given), silently wiping descriptions/bin locations/extra columns on
 * every part that wasn't having that particular field changed.
 *
 * Two independent halves, since they touch entirely different stores:
 * `updates` (per-part source field writes, run in parallel — a part
 * with no writable source or a failed write is reported in `results`
 * rather than failing the whole batch) and `addTag`/`removeTag`
 * (applied to every part_number in `partNumbers` regardless of source,
 * since tags are app-managed — see lib/tags-store.ts — and never touch
 * a SharePoint workbook at all).
 */
export async function POST(request: Request) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const body = await request.json().catch(() => null);
  if (!body || typeof body !== "object") {
    return NextResponse.json({ error: "Invalid request body." }, { status: 400 });
  }
  const record = body as Record<string, unknown>;

  const rawUpdates = Array.isArray(record.updates) ? record.updates : [];
  const updates: BulkFieldUpdate[] = [];
  for (const raw of rawUpdates) {
    if (
      raw &&
      typeof raw === "object" &&
      typeof (raw as Record<string, unknown>).id === "string" &&
      typeof (raw as Record<string, unknown>).sourceId === "string" &&
      typeof (raw as Record<string, unknown>).existingPartNumber === "string" &&
      isValidFields((raw as Record<string, unknown>).fields)
    ) {
      const r = raw as Record<string, unknown>;
      const occurrence = Number(r.existingPartOccurrence);
      updates.push({
        id: r.id as string,
        sourceId: r.sourceId as string,
        existingPartNumber: r.existingPartNumber as string,
        existingPartOccurrence: Number.isFinite(occurrence) && occurrence >= 1 ? Math.trunc(occurrence) : 1,
        fields: r.fields as PartFields,
      });
    }
  }

  const partNumbers = Array.isArray(record.partNumbers)
    ? record.partNumbers.filter((v): v is string => typeof v === "string")
    : [];
  const addTag =
    typeof record.addTag === "string" && record.addTag.trim() ? record.addTag.trim() : null;
  const removeTag =
    typeof record.removeTag === "string" && record.removeTag.trim()
      ? record.removeTag.trim()
      : null;

  const results: BulkResult[] = [];

  if (updates.length > 0) {
    const accessToken = session.accessToken;
    if (!accessToken) {
      return NextResponse.json(
        { error: "Signed-in session has no Microsoft access token yet — sign out and back in." },
        { status: 401 },
      );
    }

    await Promise.all(
      updates.map(async (update) => {
        try {
          const resolved = await resolveSharePointConfigById(update.sourceId);
          if (!resolved) {
            results.push({
              id: update.id,
              part_number: update.existingPartNumber,
              ok: false,
              error: "That source no longer exists.",
            });
            return;
          }
          const { client, fileBase } = await connectToWorkbook(accessToken, resolved.config);
          await updatePartInWorkbook(
            client,
            fileBase,
            resolved.config.tableName,
            resolved.config.columnMap ?? COLUMN_MAP,
            update.existingPartNumber,
            update.fields,
            update.existingPartOccurrence,
          );
          results.push({ id: update.id, part_number: update.existingPartNumber, ok: true });
        } catch (error) {
          results.push({
            id: update.id,
            part_number: update.existingPartNumber,
            ok: false,
            error: describeError(error),
          });
        }
      }),
    );
  }

  if ((addTag || removeTag) && partNumbers.length > 0) {
    const currentByPart = await getTagsForParts(partNumbers);
    await Promise.all(
      partNumbers.map(async (partNumber) => {
        const current = currentByPart[partNumber] ?? [];
        let next = current;
        if (addTag && !next.includes(addTag)) next = [...next, addTag];
        if (removeTag) next = next.filter((tag) => tag !== removeTag);
        if (next !== current) {
          await setTagsForPart(partNumber, next);
        }
      }),
    );
  }

  return NextResponse.json({ results });
}
