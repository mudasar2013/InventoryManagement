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
 * sources report the same part_number with different fields (a different
 * quantity_on_hand, say), the field values come from whichever source is
 * ranked highest; every source that reported the part_number is still
 * recorded in `sourceIds` so a conflict is visible rather than silently
 * dropped.
 *
 * The same source can itself report the same part_number more than once —
 * a real sheet ends up with this when a part number is reused for a later
 * restock (a new row appended after the part had previously sold through
 * and its old row was left in place, marked out of inventory). Ties are
 * broken by *last* occurrence rather than first, so a freshly appended row
 * — which is how both addPartToWorkbook write paths add a new part —
 * shadows a stale earlier row for the same part_number instead of being
 * silently hidden behind it.
 */
export function mergeParts(results: SourceParts[], priority: string[]): Part[] {
  const rank = new Map(priority.map((id, index) => [id, index]));
  const byPartNumber = new Map<
    string,
    { raw: RawPart; rank: number; sourceIds: string[] }
  >();

  for (const { sourceId, parts } of results) {
    const sourceRank = rank.get(sourceId) ?? Number.MAX_SAFE_INTEGER;
    for (const raw of parts) {
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
      if (sourceRank <= existing.rank) {
        existing.raw = raw;
        existing.rank = sourceRank;
      }
    }
  }

  return Array.from(byPartNumber.values()).map(({ raw, sourceIds }) => ({
    ...toPart(raw),
    sourceIds,
  }));
}
