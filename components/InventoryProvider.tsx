"use client";

import { createContext, useContext, useMemo, useState } from "react";
import {
  getJobById,
  getJobsForPart,
  getPartById,
  isPartLinkedToJob,
  linkPartToJob,
} from "@/lib/inventory";
import type { SourceStatus } from "@/lib/getInventory";
import { deriveStatus } from "@/lib/status";
import type { ExtraFields, Job, JobPart, Part } from "@/lib/types";

type LinkResult =
  | { ok: true; alreadyLinked: boolean; part: Part; job: Job }
  | { ok: false; reason: "missing-part" | "missing-job" };

/** A source a part can be written to — everything the "Add a part" /
 *  "Edit part" forms need for a picker, and nothing else (never the
 *  full SourceStatus, which carries error detail that isn't this
 *  context's concern). Excludes the local demo catalog (id "local"),
 *  which isn't a SharePoint workbook and can't be written to. */
export type WritableSourceOption = { id: string; label: string };

/** The fields a technician can edit or supply for a part — mirrors
 *  lib/sources/sharepoint-excel-source.ts's PartFields, kept as a
 *  separate type since this module is client-only and that one pulls
 *  in the Graph SDK. `category` and `extraFields` are optional since a
 *  source without those columns configured never reports them. `tags`
 *  is always present (defaulting to []) since it's app-managed, not
 *  read from any source. */
export type EditablePartFields = {
  part_number: string;
  description: string;
  bin_location: string;
  quantity_on_hand: number;
  category?: string;
  extraFields?: ExtraFields;
  tags?: string[];
};

function slugify(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
}

type InventoryContextValue = {
  parts: Part[];
  jobs: Job[];
  jobParts: JobPart[];
  lastLinkMessage: string | null;
  linkPart: (jobId: string, partId: string) => LinkResult;
  jobsForPart: (partId: string) => Job[];
  isLinked: (jobId: string, partId: string) => boolean;
  /** SharePoint sources a part can be added to or edited in — see
   *  WritableSourceOption. */
  writableSources: WritableSourceOption[];
  /** The full app-managed tag vocabulary (see lib/tags-store.ts) — for
   *  the tag filter on the Parts page and the tag picker on the part
   *  detail edit form. Managed from the Settings page, which reloads
   *  via router.refresh() rather than a local-state method here. */
  tags: string[];
  /** Applies a successful PATCH /api/parts write to local state — the
   *  API call already wrote the source workbook; this just keeps the
   *  in-memory catalog in sync without a full reload. `previousId` is
   *  the part's id before the edit, since a part_number rename changes
   *  its synthesized id too — returns the part's new id so a caller can
   *  navigate to it. */
  applyPartUpdate: (previousId: string, fields: EditablePartFields) => string;
  /** Applies a successful POST /api/parts (add) to local state — see
   *  applyPartUpdate. Returns the new part's id. */
  applyNewPart: (fields: EditablePartFields, sourceId: string) => string;
  /** Applies a successful POST /api/parts/bulk write to local state —
   *  one partial field set per part_number (only the fields that were
   *  actually part of the bulk edit; anything a given part's source
   *  couldn't write is simply absent from its entry, same as a normal
   *  edit skipping an unmapped column). Keyed by part_number rather
   *  than id since a bulk edit never renames parts. */
  applyBulkUpdate: (
    updates: { part_number: string; fields: Partial<EditablePartFields> }[],
  ) => void;
};

const InventoryContext = createContext<InventoryContextValue | null>(null);

export function InventoryProvider({
  initialParts,
  initialJobs,
  initialJobParts,
  initialSourceStatuses,
  initialTags,
  children,
}: {
  initialParts: Part[];
  initialJobs: Job[];
  initialJobParts: JobPart[];
  initialSourceStatuses: SourceStatus[];
  initialTags: string[];
  children: React.ReactNode;
}) {
  // Seeded from the server's loadInventory() result (see app/layout.tsx)
  // instead of importing mock data directly, so the client never has its
  // own disconnected copy of the catalog.
  const [parts, setParts] = useState<Part[]>(() => initialParts.map((part) => ({ ...part })));
  const [jobs] = useState<Job[]>(() => initialJobs);
  const [jobParts, setJobParts] = useState<JobPart[]>(() =>
    initialJobParts.map((item) => ({ ...item })),
  );
  const [lastLinkMessage, setLastLinkMessage] = useState<string | null>(null);
  const [writableSources] = useState<WritableSourceOption[]>(() =>
    initialSourceStatuses
      .filter((status) => status.id !== "local")
      .map((status) => ({ id: status.id, label: status.label })),
  );
  const [tags] = useState<string[]>(() => initialTags);

  const value = useMemo<InventoryContextValue>(() => {
    return {
      parts,
      jobs,
      jobParts,
      lastLinkMessage,
      writableSources,
      tags,
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
      applyPartUpdate: (previousId: string, fields: EditablePartFields) => {
        const newId = `sharepoint-${slugify(fields.part_number)}`;
        setParts((current) =>
          current.map((part) =>
            part.id === previousId
              ? {
                  ...part,
                  ...fields,
                  // A part_number rename changes its synthesized id too —
                  // see the matching logic in mapTableRowsToRawParts. A
                  // sheet with its own id column would keep the same id
                  // across a rename, but this app doesn't know a part's id
                  // came from one vs. was synthesized, so it re-derives it
                  // the same way a fresh fetch would for the common case.
                  id: newId,
                  status: deriveStatus(fields.quantity_on_hand),
                }
              : part,
          ),
        );
        return newId;
      },
      applyNewPart: (fields: EditablePartFields, sourceId: string) => {
        const newId = `sharepoint-${slugify(fields.part_number)}`;
        setParts((current) => [
          ...current,
          {
            id: newId,
            ...fields,
            status: deriveStatus(fields.quantity_on_hand),
            sourceIds: [sourceId],
          },
        ]);
        return newId;
      },
      applyBulkUpdate: (updates) => {
        const byPartNumber = new Map(updates.map((update) => [update.part_number, update.fields]));
        setParts((current) =>
          current.map((part) => {
            const fields = byPartNumber.get(part.part_number);
            if (!fields) return part;
            return {
              ...part,
              ...fields,
              status:
                fields.quantity_on_hand !== undefined
                  ? deriveStatus(fields.quantity_on_hand)
                  : part.status,
            };
          }),
        );
      },
    };
  }, [jobs, jobParts, lastLinkMessage, parts, tags, writableSources]);

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
