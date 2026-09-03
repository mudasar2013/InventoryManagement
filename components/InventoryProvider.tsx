"use client";

import { createContext, useContext, useMemo, useState } from "react";
import {
  getJobById,
  getJobsForPart,
  getPartById,
  isPartLinkedToJob,
  linkPartToJob,
} from "@/lib/inventory";
import type { Job, JobPart, Part } from "@/lib/types";

type LinkResult =
  | { ok: true; alreadyLinked: boolean; part: Part; job: Job }
  | { ok: false; reason: "missing-part" | "missing-job" };

type InventoryContextValue = {
  parts: Part[];
  jobs: Job[];
  jobParts: JobPart[];
  lastLinkMessage: string | null;
  linkPart: (jobId: string, partId: string) => LinkResult;
  jobsForPart: (partId: string) => Job[];
  isLinked: (jobId: string, partId: string) => boolean;
};

const InventoryContext = createContext<InventoryContextValue | null>(null);

export function InventoryProvider({
  initialParts,
  initialJobs,
  initialJobParts,
  children,
}: {
  initialParts: Part[];
  initialJobs: Job[];
  initialJobParts: JobPart[];
  children: React.ReactNode;
}) {
  // Seeded from the server's loadInventory() result (see app/layout.tsx)
  // instead of importing mock data directly, so the client never has its
  // own disconnected copy of the catalog.
  const [parts] = useState<Part[]>(() => initialParts.map((part) => ({ ...part })));
  const [jobs] = useState<Job[]>(() => initialJobs);
  const [jobParts, setJobParts] = useState<JobPart[]>(() =>
    initialJobParts.map((item) => ({ ...item })),
  );
  const [lastLinkMessage, setLastLinkMessage] = useState<string | null>(null);

  const value = useMemo<InventoryContextValue>(() => {
    return {
      parts,
      jobs,
      jobParts,
      lastLinkMessage,
      jobsForPart: (partId: string) => getJobsForPart(partId, jobParts, jobs),
      isLinked: (jobId: string, partId: string) =>
        isPartLinkedToJob(jobParts, jobId, partId),
      linkPart: (jobId: string, partId: string) => {
        const part = getPartById(partId, parts);
        const job = getJobById(jobId, jobs);
        if (!part) return { ok: false, reason: "missing-part" };
        if (!job) return { ok: false, reason: "missing-job" };

        const alreadyLinked = isPartLinkedToJob(jobParts, jobId, partId);
        if (!alreadyLinked) {
          setJobParts((current) => linkPartToJob(current, jobId, partId));
        }

        const quantity = part.quantity_on_hand;
        setLastLinkMessage(
          alreadyLinked
            ? `${part.part_number} is already on job ${job.job_number}. Quantity on hand is still ${quantity}.`
            : `Linked ${part.part_number} to job ${job.job_number}. Quantity on hand is still ${quantity} — attaching does not deduct stock.`,
        );

        return { ok: true, alreadyLinked, part, job };
      },
    };
  }, [jobs, jobParts, lastLinkMessage, parts]);

  return (
    <InventoryContext.Provider value={value}>{children}</InventoryContext.Provider>
  );
}

export function useInventory() {
  const context = useContext(InventoryContext);
  if (!context) {
    throw new Error("useInventory must be used inside InventoryProvider");
  }
  return context;
}
