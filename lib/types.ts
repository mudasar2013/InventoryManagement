export type PartStatus = "In Stock" | "Low Stock" | "Out of Stock";

/** One column from a source's sheet that isn't one of the named
 *  RawPart fields (part_number, description, bin_location,
 *  quantity_on_hand, category) — captured so the part detail page can
 *  show (and edit) every column a workbook has, not just the ones this
 *  app has a dedicated field for. `kind` is inferred from the sheet's
 *  own header text and values (see sharepoint-excel-source.ts for the
 *  exact rules):
 *  - "boolean" — the header names itself "(Yes/No)", or every non-blank
 *    cell in the column is some spelling of yes/no. Rendered as a
 *    checkbox / Yes-No dropdown.
 *  - "date" — the header contains "date" (e.g. "Entry Date", "Date
 *    Rcvd"). The sheet stores these as raw Excel serial numbers;
 *    `value` here is always a normalized "YYYY-MM-DD" string (or "" for
 *    blank), converted back to a serial only when writing.
 *  - "select" — a small set of headers this app knows the shop's fixed
 *    vocabulary for (currently just "Condition") — `options` lists the
 *    fixed choices to offer, but `value` can still be free text when
 *    the sheet already has something outside that list.
 *  - "text" — anything else. */
export type ExtraFieldKind = "text" | "boolean" | "date" | "select";
export interface ExtraField {
  kind: ExtraFieldKind;
  value: string | boolean;
  /** Only meaningful for kind "select" — the fixed choices to offer
   *  (see SELECT_FIELD_OPTIONS in sharepoint-excel-source.ts). A value
   *  that matches none of these is still preserved and shown as
   *  free text ("Other") rather than silently discarded. */
  options?: string[];
}
export type ExtraFields = Record<string, ExtraField>;

/** Fixed dropdown choices for the "Condition" extra field. Shared by
 *  the server-side classification (sharepoint-excel-source.ts, which
 *  reports these as the "select" kind's options) and the client-side
 *  "Add a part" form (AddPartForm.tsx, which lets a technician set
 *  Condition on a brand-new part) — kept in one place so the two lists
 *  can't drift apart. "Other" is always last: it's the free-text
 *  escape hatch, not a real condition. */
export const CONDITION_OPTIONS = ["New", "Used", "OpenBox", "Used/Working", "Other"] as const;

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
