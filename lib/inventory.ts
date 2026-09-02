import { jobParts as seedJobParts, jobs, parts } from "./mockData";
import type { Job, JobPart, Part, PartStatus } from "./types";

export function getPartById(id: string, catalog: Part[] = parts): Part | undefined {
  return catalog.find((part) => part.id === id);
}

export function getJobById(id: string): Job | undefined {
  return jobs.find((job) => job.id === id);
}

export function getJobPartsForJob(jobId: string, links: JobPart[] = seedJobParts): JobPart[] {
  return links.filter((item) => item.job_id === jobId);
}

export function getJobsForPart(partId: string, links: JobPart[] = seedJobParts): Job[] {
  const jobIds = new Set(
    links.filter((item) => item.part_id === partId).map((item) => item.job_id),
  );
  return jobs.filter((job) => jobIds.has(job.id));
}

export function searchParts(
  catalog: Part[],
  query: string,
  status?: PartStatus | "All",
): Part[] {
  const normalized = query.trim().toLowerCase();

  return catalog.filter((part) => {
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

export function countByStatus(catalog: Part[] = parts) {
  return {
    all: catalog.length,
    inStock: catalog.filter((part) => part.status === "In Stock").length,
    lowStock: catalog.filter((part) => part.status === "Low Stock").length,
    outOfStock: catalog.filter((part) => part.status === "Out of Stock").length,
  };
}

/**
 * Link a catalog part to a job. This only writes a JobPart row.
 * Quantity on hand is never read or written here — attaching is reservation-only.
 */
export function linkPartToJob(
  links: JobPart[],
  jobId: string,
  partId: string,
  quantityNeeded = 1,
): JobPart[] {
  const alreadyLinked = links.some(
    (item) => item.job_id === jobId && item.part_id === partId,
  );
  if (alreadyLinked) {
    return links;
  }

  return [
    ...links,
    {
      id: `jp-${jobId}-${partId}`,
      job_id: jobId,
      part_id: partId,
      quantity_needed: quantityNeeded,
    },
  ];
}

export function isPartLinkedToJob(
  links: JobPart[],
  jobId: string,
  partId: string,
): boolean {
  return links.some((item) => item.job_id === jobId && item.part_id === partId);
}
