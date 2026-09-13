export type PartStatus = "In Stock" | "Low Stock" | "Out of Stock";

/** One column from a source's sheet that isn't one of the named
 *  RawPart fields (part_number, description, bin_location,
 *  quantity_on_hand, category) — captured so the part detail page can
 *  show (and edit) every column a workbook has, not just the ones this
 *  app has a dedicated field for. `kind` is inferred from the sheet's
 *  own values: a column where every non-blank cell is some spelling of
 *  Yes/No is "boolean" (rendered as a checkbox); anything else is
 *  "text". See columnLooksBoolean in sharepoint-excel-source.ts. */
export type ExtraFieldKind = "text" | "boolean";
export interface ExtraField {
  kind: ExtraFieldKind;
  value: string | boolean;
}
export type ExtraFields = Record<string, ExtraField>;

/**
 * The fields an inventory source is expected to supply for a part.
 * Deliberately excludes `status` — status is a derived value (see
 * lib/status.ts), never something a source authors directly. Sources
 * disagree on thresholds and go stale; deriving it from quantity_on_hand
 * at read time is the only way it stays correct across sources.
 */
export interface RawPart {
  id: string;
  part_number: string;
  description: string;
  bin_location: string;
  quantity_on_hand: number;
  /** From the source's own "Category" column, when one is mapped (see
   *  SharePointColumnMap.category) — undefined when this source has no
   *  category column configured. Filtering by category (see
   *  PartsExplorer) reads whatever distinct values show up here rather
   *  than a separately curated list. */
  category?: string;
  /** Every other column this source's sheet has, keyed by that column's
   *  own header text — see ExtraField. Always present (possibly empty)
   *  once a part has been through mapTableRowsToRawParts; absent only
   *  on hand-built fixtures that don't set it. */
  extraFields?: ExtraFields;
}

export interface Part extends RawPart {
  status: PartStatus;
  /**
   * Which inventory source(s) reported this part_number, in case of a
   * conflict across sources. Populated by lib/sources/merge.ts. Optional
   * because hand-built Part fixtures (tests, mock UI data) don't need it.
   */
  sourceIds?: string[];
  /** App-managed tags (see lib/tags-store.ts) — never read from or
   *  written to a source's sheet, unlike every other Part field. Keyed
   *  by part_number in Redis, so a tag survives a part being re-read
   *  from its source (or reported by a different source entirely).
   *  Always an array (possibly empty) once attached by loadInventory. */
  tags?: string[];
}

export type JobStatus = "Scheduled" | "In Progress" | "Completed";

export interface Job {
  id: string;
  job_number: string;
  customer_name: string;
  appliance: string;
  issue: string;
  status: JobStatus;
  scheduled_for: string;
}

export interface JobPart {
  id: string;
  job_id: string;
  part_id: string;
  quantity_needed: number;
}
