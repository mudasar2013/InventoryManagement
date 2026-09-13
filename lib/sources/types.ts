import type { Job, JobPart, RawPart } from "../types";

/**
 * One place inventory data can come from. `id` must be stable and unique
 * across the sources an app is configured with — it's what merge.ts uses
 * to record provenance and resolve conflicts, and what loadInventory()
 * uses as the priority key.
 *
 * A future SharePoint source implements exactly this: read a list or an
 * Excel workbook via Microsoft Graph and map each row into a RawPart.
 * Nothing else in the app needs to change — loadInventory() and the merge
 * step are already written against this interface.
 */
export interface InventorySource {
  id: string;
  label: string;
  fetchParts(): Promise<RawPart[]>;
  fetchJobs(): Promise<Job[]>;
  fetchJobParts(): Promise<JobPart[]>;
  /** A browser URL for the underlying file/sheet this source reads
   *  from, e.g. so a technician can open a SharePoint workbook directly
   *  to make a change by hand — see the "Data sources" page. Optional:
   *  not every source has an underlying file (the local catalog
   *  doesn't), and for those that do it may only be known after
   *  fetchParts() has run at least once this request. */
  getFileUrl?(): string | undefined;
}
