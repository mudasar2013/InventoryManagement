import { jobParts, jobs, parts } from "./mockData";
import type { Job, JobPart, Part, PartStatus } from "./types";

export function getPartById(id: string): Part | undefined {
  return parts.find((part) => part.id === id);
}

export function getJobById(id: string): Job | undefined {
  return jobs.find((job) => job.id === id);
}

export function getJobPartsForJob(jobId: string): JobPart[] {
  return jobParts.filter((item) => item.job_id === jobId);
}

export function getJobsForPart(partId: string): Job[] {
  const jobIds = new Set(
    jobParts.filter((item) => item.part_id === partId).map((item) => item.job_id),
  );
  return jobs.filter((job) => jobIds.has(job.id));
}

export function searchParts(query: string, status?: PartStatus | "All"): Part[] {
  const normalized = query.trim().toLowerCase();

  return parts.filter((part) => {
    const matchesStatus = !status || status === "All" || part.status === status;
    if (!matchesStatus) {
      return false;
    }

    if (!normalized) {
      return true;
    }

    return (
      part.part_number.toLowerCase().includes(normalized) ||
      part.description.toLowerCase().includes(normalized) ||
      part.bin_location.toLowerCase().includes(normalized)
    );
  });
}

export function countByStatus() {
  return {
    all: parts.length,
    inStock: parts.filter((part) => part.status === "In Stock").length,
    lowStock: parts.filter((part) => part.status === "Low Stock").length,
    outOfStock: parts.filter((part) => part.status === "Out of Stock").length,
  };
}
