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
}
