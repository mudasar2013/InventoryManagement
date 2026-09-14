import { toPart } from "../status";
import type { Part, RawPart } from "../types";

export interface SourceParts {
  sourceId: string;
  parts: RawPart[];
}

/**
 * Combines every source's reported parts into one flat list, tagging
 * each with the single source it came from.
 *
 * This deliberately does NOT merge rows across different sources, even
 * when two sources report the identical part_number. "Hadi Inventory"
 * and "7300 Inventory" reporting the same part_number are two distinct
 * physical stock records — different bins, different quantities on
 * hand, possibly different conditions — not one part described two
 * different ways. Folding them into a single card silently threw away
 * one source's bin location and quantity and made it impossible to
 * open, edit, or count either row on its own. So every row from every
 * source becomes its own Part, exactly as a source repeating the same
 * part_number within itself already did (see RawPart.partNumberOccurrence
 * and mapTableRowsToRawParts in sharepoint-excel-source.ts, which is
 * also what keeps every row's `id` unique across sources now that none
 * of them are combined).
 *
 * `sourceIds` on the returned Part is always a single-element array —
 * kept as an array (rather than a lone `sourceId` field) so existing
 * code that reads `part.sourceIds` (the source-hint badge on
 * PartCard/PartDetail, the Reports "Source" column) doesn't need a
 * separate single-vs-multi code path.
 */
export function mergeParts(results: SourceParts[]): Part[] {
  const parts: Part[] = [];
  for (const { sourceId, parts: rawParts } of results) {
    for (const raw of rawParts) {
      parts.push({ ...toPart(raw), sourceIds: [sourceId] });
    }
  }
  return parts;
}
