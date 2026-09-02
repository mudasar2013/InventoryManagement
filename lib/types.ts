export type PartStatus = "In Stock" | "Low Stock" | "Out of Stock";

export interface Part {
  id: string;
  part_number: string;
  description: string;
  bin_location: string;
  quantity_on_hand: number;
  status: PartStatus;
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
