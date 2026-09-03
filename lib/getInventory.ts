import { cache } from "react";
import { describeError } from "./describe-error";
import { localSource } from "./sources/local-source";
import { mergeParts } from "./sources/merge";
import {
  createSharePointExcelSource,
  readSharePointExcelConfig,
} from "./sources/sharepoint-excel-source";
import type { InventorySource } from "./sources/types";
import type { Job, JobPart, Part } from "./types";

export interface Inventory {
  parts: Part[];
  jobs: Job[];
  jobParts: JobPart[];
  /** Human-readable notices when a source failed and was skipped, so the
   *  UI can say "SharePoint is unavailable" instead of silently showing
   *  an incomplete catalog with no explanation. */
  warnings: string[];
}

function buildSources(accessToken: string | undefined): {
  sources: InventorySource[];
  priority: string[];
} {
  const sources: InventorySource[] = [localSource];
  // SharePoint outranks the bundled demo catalog when both report the
  // same part_number — see lib/sources/merge.ts.
  const priority = ["sharepoint", "local"];

  const sharePointConfig = readSharePointExcelConfig();
  if (sharePointConfig && accessToken) {
    sources.unshift(createSharePointExcelSource(accessToken, sharePointConfig));
  }

  return { sources, priority };
}

/**
 * The single place that turns configured sources into the data every
 * page and component renders. Both the server-side existence check in
 * app/parts/[id]/page.tsx and the client InventoryProvider are seeded
 * from this same call, so they can no longer disagree about what's in
 * the catalog.
 *
 * `accessToken` is the signed-in technician's Graph token (from
 * getServerSession — see app/layout.tsx). Passed as a plain argument
 * rather than buried in an options object so React's cache() can
 * actually dedupe repeat calls with the same token within one request.
 *
 * A source failing (SharePoint unreachable, token expired, a workbook
 * column renamed) does not take down the page — it's dropped and
 * recorded in `warnings` so the rest of the catalog still renders.
 */
export const loadInventory = cache(
  async (accessToken?: string): Promise<Inventory> => {
    const { sources, priority } = buildSources(accessToken);
    const warnings: string[] = [];

    const results = await Promise.all(
      sources.map(async (source) => {
        try {
          const [parts, jobs, jobParts] = await Promise.all([
            source.fetchParts(),
            source.fetchJobs(),
            source.fetchJobParts(),
          ]);
          return { sourceId: source.id, parts, jobs, jobParts };
        } catch (error) {
          warnings.push(
            `${source.label} is unavailable right now (${describeError(
              error,
            )}). Showing the rest of the catalog.`,
          );
          return { sourceId: source.id, parts: [], jobs: [], jobParts: [] };
        }
      }),
    );

    const parts = mergeParts(
      results.map(({ sourceId, parts }) => ({ sourceId, parts })),
      priority,
    );

    return {
      parts,
      jobs: results.flatMap((result) => result.jobs),
      jobParts: results.flatMap((result) => result.jobParts),
      warnings,
    };
  },
);
