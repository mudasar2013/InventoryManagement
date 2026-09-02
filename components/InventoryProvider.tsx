"use client";

import { createContext, useContext, useMemo, useState } from "react";
import {
  getJobById,
  getJobsForPart,
  getPartById,
  isPartLinkedToJob,
  linkPartToJob,
} from "@/lib/inventory";
import { jobParts as seedJobParts, jobs as seedJobs, parts as seedParts } from "@/lib/mockData";
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

export function InventoryProvider({ children }: { children: React.ReactNode }) {
  const [parts] = useState<Part[]>(() => seedParts.map((part) => ({ ...part })));
  const [jobParts, setJobParts] = useState<JobPart[]>(() =>
    seedJobParts.map((item) => ({ ...item })),
  );
  const [lastLinkMessage, setLastLinkMessage] = useState<string | null>(null);

  const value = useMemo<InventoryContextValue>(() => {
    return {
      parts,
      jobs: seedJobs,
      jobParts,
      lastLinkMessage,
      jobsForPart: (partId: string) => getJobsForPart(partId, jobParts),
      isLinked: (jobId: string, partId: string) =>
        isPartLinkedToJob(jobParts, jobId, partId),
      linkPart: (jobId: string, partId: string) => {
        const part = getPartById(partId, parts);
        const job = getJobById(jobId);
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
  }, [jobParts, lastLinkMessage, parts]);

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
