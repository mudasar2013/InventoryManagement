export type PartStatus = "In Stock" | "Low Stock" | "Out of Stock";

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
}

export interface Part extends RawPart {
  status: PartStatus;
  /**
   * Which inventory source(s) reported this part_number, in case of a
   * conflict across sources. Populated by lib/sources/merge.ts. Optional
   * because hand-built Part fixtures (tests, mock UI data) don't need it.
   */
  sourceIds?: string[];
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
