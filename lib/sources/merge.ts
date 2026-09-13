import { toPart } from "../status";
import type { Part, RawPart } from "../types";

export interface SourceParts {
  sourceId: string;
  parts: RawPart[];
}

/**
 * Merge parts reported by multiple sources, keyed by part_number — the
 * real-world identifier a technician searches by — not by each source's
 * internal row id, which will differ between systems.
 *
 * `priority` is a list of source ids, highest priority first. When two
 * DIFFERENT sources report the same part_number, the field values come
 * from whichever source is ranked highest; every source that reported
 * the part_number is still recorded in `sourceIds` so a conflict is
 * visible rather than silently dropped.
 *
 * The SAME source can itself report the same part_number more than
 * once — a real sheet ends up with this when a part number is reused
 * for a later restock (a new row appended after the part had
 * previously sold through and its old row was left in place, marked
 * out of inventory). That isn't the same conflict a cross-source
 * mismatch is: it's two genuinely different physical rows, not one
 * part described two different ways — so neither shadows the other.
 * Only a source's *first* row for a given part_number is eligible to
 * combine with another source's report of that part_number; every
 * later row that source reports for the same part_number stands
 * entirely on its own as a separate Part (mapTableRowsToRawParts gives
 * each such row its own id, so this never collides with the first
 * one's).
 */
export function mergeParts(results: SourceParts[], priority: string[]): Part[] {
  const rank = new Map(priority.map((id, index) => [id, index]));
  const byPartNumber = new Map<
    string,
    { raw: RawPart; rank: number; sourceIds: string[] }
  >();
  const standalone: { raw: RawPart; sourceIds: string[] }[] = [];

  for (const { sourceId, parts } of results) {
    const sourceRank = rank.get(sourceId) ?? Number.MAX_SAFE_INTEGER;
    const alreadyReportedBySource = new Set<string>();

    for (const raw of parts) {
      if (alreadyReportedBySource.has(raw.part_number)) {
        standalone.push({ raw, sourceIds: [sourceId] });
        continue;
      }
      alreadyReportedBySource.add(raw.part_number);

      const existing = byPartNumber.get(raw.part_number);
      if (!existing) {
        byPartNumber.set(raw.part_number, {
          raw,
          rank: sourceRank,
          sourceIds: [sourceId],
        });
        continue;
      }

      existing.sourceIds.push(sourceId);
      if (sourceRank < existing.rank) {
        existing.raw = raw;
        existing.rank = sourceRank;
      }
    }
  }

  return [
    ...Array.from(byPartNumber.values()).map(({ raw, sourceIds }) => ({
      ...toPart(raw),
      sourceIds,
    })),
    ...standalone.map(({ raw, sourceIds }) => ({
      ...toPart(raw),
      sourceIds,
    })),
  ];
}
