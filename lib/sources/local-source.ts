import { jobParts, jobs, parts as rawParts } from "../mockData";
import type { InventorySource } from "./types";

/**
 * Wraps the bundled mock catalog as an InventorySource. This is the only
 * source configured today (see lib/getInventory.ts); a SharePoint source
 * (or any other) plugs in beside it without touching callers.
 */
export const localSource: InventorySource = {
  id: "local",
  label: "Local catalog",
  async fetchParts() {
    return rawParts;
  },
  async fetchJobs() {
    return jobs;
  },
  async fetchJobParts() {
    return jobParts;
  },
};
