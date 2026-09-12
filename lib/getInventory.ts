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

/**
 * Status of one configured (or configurable) inventory source, for the
 * "Data sources" settings page — see app/sources/page.tsx. Distinct from
 * `warnings` below: `warnings` is a short user-facing phrase for the main
 * catalog pages, this is the structured per-source detail that page needs
 * to render a real status list instead of one lumped-together message.
 */
export interface SourceStatus {
  id: string;
  label: string;
  /** Extra human-readable identifying detail — e.g. which site/workbook a
   *  SharePoint source points at. Never includes secrets. */
  detail?: string;
  /** Whether this source type has the configuration it needs to run at
   *  all (env vars set, etc.) — independent of whether the most recent
   *  fetch actually succeeded. A source can be configured and still fail
   *  (network, expired token, renamed column). */
  configured: boolean;
  /** Result of the most recent fetch attempt this request. Undefined when
   *  the source was never attempted — not configured, or (for SharePoint)
   *  configured but this request has no Graph access token to use yet. */
  ok?: boolean;
  /** Part count from the most recent successful fetch. */
  partCount?: number;
  /** Short reason the most recent fetch failed, or why it couldn't be
   *  attempted despite being configured. Never raw error internals — see
   *  describeError. */
  note?: string;
}

export interface Inventory {
  parts: Part[];
  jobs: Job[];
  jobParts: JobPart[];
  /** Human-readable notices when a source failed and was skipped, so the
   *  UI can say "SharePoint is unavailable" instead of silently showing
   *  an incomplete catalog with no explanation. */
  warnings: string[];
  /** Every source type the app knows about (configured or not), for the
   *  "Data sources" settings page. */
  sourceStatuses: SourceStatus[];
}

interface SharePointSourceInfo {
  configured: boolean;
  detail?: string;
}

function buildSources(accessToken: string | undefined): {
  sources: InventorySource[];
  priority: string[];
  sharePoint: SharePointSourceInfo;
} {
  const sources: InventorySource[] = [localSource];
  // SharePoint outranks the bundled demo catalog when both report the
  // same part_number — see lib/sources/merge.ts.
  const priority = ["sharepoint", "local"];

  const sharePointConfig = readSharePointExcelConfig();
  const sharePoint: SharePointSourceInfo = {
    configured: sharePointConfig !== null,
    detail: sharePointConfig
      ? `${sharePointConfig.siteHostname}${sharePointConfig.sitePath} · ${sharePointConfig.filePath} (table: ${sharePointConfig.tableName})`
      : undefined,
  };

  if (sharePointConfig && accessToken) {
    sources.unshift(createSharePointExcelSource(accessToken, sharePointConfig));
  }

  return { sources, priority, sharePoint };
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
    const { sources, priority, sharePoint } = buildSources(accessToken);
    const warnings: string[] = [];

    const results = await Promise.all(
      sources.map(async (source) => {
        try {
          const [parts, jobs, jobParts] = await Promise.all([
            source.fetchParts(),
            source.fetchJobs(),
            source.fetchJobParts(),
          ]);
          return { sourceId: source.id, parts, jobs, jobParts, ok: true as const };
        } catch (error) {
          // Full detail (including nested .cause) goes to the server
          // terminal — the warning banner only gets a short phrase, and
          // some causes (proxy/TLS internals, stack traces) are more
          // than a shop-floor user needs to see.
          console.error(`[getInventory] ${source.id} source failed:`, error);
          const note = describeError(error);
          warnings.push(
            `${source.label} is unavailable right now (${note}). Showing the rest of the catalog.`,
          );
          return {
            sourceId: source.id,
            parts: [],
            jobs: [],
            jobParts: [],
            ok: false as const,
            note,
          };
        }
      }),
    );

    const parts = mergeParts(
      results.map(({ sourceId, parts }) => ({ sourceId, parts })),
      priority,
    );

    const localResult = results.find((result) => result.sourceId === "local");
    const sharePointResult = results.find((result) => result.sourceId === "sharepoint");

    const sourceStatuses: SourceStatus[] = [
      {
        id: "local",
        label: "Local catalog",
        configured: true,
        ok: localResult?.ok,
        partCount: localResult?.parts.length,
        note: localResult && !localResult.ok ? localResult.note : undefined,
      },
      {
        id: "sharepoint",
        label: "SharePoint workbook",
        detail: sharePoint.detail,
        configured: sharePoint.configured,
        ok: sharePointResult?.ok,
        partCount: sharePointResult?.parts.length,
        note: sharePointResult
          ? sharePointResult.ok
            ? undefined
            : sharePointResult.note
          : sharePoint.configured
            ? "Signed-in session has no Microsoft access token yet — sign out and back in."
            : "Not configured. Set SHAREPOINT_SITE_HOSTNAME, SHAREPOINT_SITE_PATH, SHAREPOINT_FILE_PATH, and SHAREPOINT_TABLE_NAME.",
      },
    ];

    return {
      parts,
      jobs: results.flatMap((result) => result.jobs),
      jobParts: results.flatMap((result) => result.jobParts),
      warnings,
      sourceStatuses,
    };
  },
);
