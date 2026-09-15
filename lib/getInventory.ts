import { cache } from "react";
import { describeError, describeErrorDetail } from "./describe-error";
import { localSource } from "./sources/local-source";
import { mergeParts } from "./sources/merge";
import {
  columnMapForStoredSource,
  ENV_SHAREPOINT_SOURCE_ID,
} from "./sources/resolve-sharepoint-source";
import { createSharePointExcelSource, readSharePointExcelConfig } from "./sources/sharepoint-excel-source";
import { isSourceStoreConfigured, listStoredSharePointSources } from "./sources/sharepoint-source-store";
import type { InventorySource } from "./sources/types";
import { getTagsForParts, isTagStoreConfigured, listTags } from "./tags-store";
import type { Job, JobPart, Part } from "./types";

/**
 * Status of one configured (or configurable) inventory source, for the
 * "Data sources" settings page — see app/sources/page.tsx. Distinct from
 * `warnings` below: `warnings` is a short user-facing phrase for the main
 * catalog pages, this is the structured per-source detail that page needs
 * to render a real status list (and an add/remove UI) instead of one
 * lumped-together message.
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
  /** Full diagnostic text for the most recent fetch failure — error
   *  name/message, the full .cause chain, and (for Microsoft Graph
   *  errors) status code/error code/request id/response body. Shown
   *  behind a "Show details" toggle on the Data sources page rather
   *  than always-on, since it's meant for pinning down a typo'd
   *  hostname/path rather than everyday reading. See describeErrorDetail. */
  debugDetail?: string;
  /** Browser URL for the underlying file, when known — see
   *  InventorySource.getFileUrl. Populated once the site + file lookup
   *  succeeds, even if the fetch goes on to fail on the table/worksheet
   *  name, so a technician can open the file to check it regardless of
   *  whether this request's fetch fully succeeded. */
  fileUrl?: string;
  /** Whether this entry can be deleted from the "Data sources" page. The
   *  legacy env-var SharePoint source and the always-on local catalog are
   *  not — removing those means editing environment variables/code. */
  removable: boolean;
}

export interface Inventory {
  parts: Part[];
  jobs: Job[];
  jobParts: JobPart[];
  /** Human-readable notices when a source failed and was skipped, so the
   *  UI can say "SharePoint is unavailable" instead of silently showing
   *  an incomplete catalog with no explanation. */
  warnings: string[];
  /** Every source currently configured (env-based or added from the
   *  page), for the "Data sources" settings page. */
  sourceStatuses: SourceStatus[];
  /** Whether the "add a SharePoint source" form should be usable — false
   *  when no Redis store is reachable to save new entries to. */
  canAddSharePointSource: boolean;
  /** The full tag vocabulary (see lib/tags-store.ts), for the Settings
   *  page and for populating the Parts page's tag filter. */
  tags: string[];
  /** Whether the Settings page's tag management (add/rename/delete) is
   *  usable — same Redis-reachability check as canAddSharePointSource. */
  canManageTags: boolean;
}

interface SharePointEntry {
  id: string;
  label: string;
  detail: string;
  removable: boolean;
}

async function buildSources(accessToken: string | undefined): Promise<{
  sources: InventorySource[];
  sharePointEntries: SharePointEntry[];
}> {
  const sources: InventorySource[] = [localSource];
  const sharePointEntries: SharePointEntry[] = [];

  // The original single SharePoint source, still configured via
  // environment variables. Kept working for whoever set this up before
  // per-source storage existed — it just can't be edited or removed from
  // the page, only by changing env vars and redeploying.
  const envConfig = readSharePointExcelConfig();
  if (envConfig) {
    const id = ENV_SHAREPOINT_SOURCE_ID;
    sharePointEntries.push({
      id,
      label: "SharePoint workbook (env)",
      detail: `${envConfig.siteHostname}${envConfig.sitePath} · ${envConfig.filePath} (table/sheet: ${envConfig.tableName})`,
      removable: false,
    });
    if (accessToken) {
      sources.unshift(
        createSharePointExcelSource(accessToken, envConfig, {
          id,
          label: "SharePoint workbook (env)",
        }),
      );
    }
  }

  // Additional SharePoint workbooks added from the "Data sources" page —
  // see lib/sources/sharepoint-source-store.ts.
  const stored = await listStoredSharePointSources();
  for (const entry of stored) {
    sharePointEntries.push({
      id: entry.id,
      label: entry.label,
      detail: `${entry.siteHostname}${entry.sitePath} · ${entry.filePath} (table/sheet: ${entry.tableName})`,
      removable: true,
    });
    if (accessToken) {
      sources.unshift(
        createSharePointExcelSource(
          accessToken,
          {
            siteHostname: entry.siteHostname,
            sitePath: entry.sitePath,
            filePath: entry.filePath,
            tableName: entry.tableName,
            columnMap: columnMapForStoredSource(entry),
          },
          { id: entry.id, label: entry.label },
        ),
      );
    }
  }

  return { sources, sharePointEntries };
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
    const { sources, sharePointEntries } = await buildSources(accessToken);
    const warnings: string[] = [];

    const results = await Promise.all(
      sources.map(async (source) => {
        try {
          const [parts, jobs, jobParts] = await Promise.all([
            source.fetchParts(),
            source.fetchJobs(),
            source.fetchJobParts(),
          ]);
          return {
            sourceId: source.id,
            parts,
            jobs,
            jobParts,
            ok: true as const,
            // Some sources (SharePoint) resolve their underlying file
            // before reading the table/worksheet — read this after the
            // fetch either way, not just on success, so a technician
            // can still open the file when the fetch fails later on a
            // bad table/worksheet name (see the catch branch below).
            fileUrl: source.getFileUrl?.(),
          };
        } catch (error) {
          // Full detail (including nested .cause) goes to the server
          // terminal — the warning banner only gets a short phrase, and
          // some causes (proxy/TLS internals, stack traces) are more
          // than a shop-floor user needs to see.
          console.error(`[getInventory] ${source.id} source failed:`, error);
          const note = describeError(error);
          const debugDetail = describeErrorDetail(error);
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
            debugDetail,
            fileUrl: source.getFileUrl?.(),
          };
        }
      }),
    );

    const mergedParts = mergeParts(results.map(({ sourceId, parts }) => ({ sourceId, parts })));

    // Tags never come from a source — attach them in one batched Redis
    // read rather than per-part, so a catalog of hundreds of parts
    // doesn't mean hundreds of lookups. A part with no tags simply
    // isn't in tagsByPartNumber; `?? []` below is what makes `tags`
    // always an array rather than sometimes undefined.
    const [tagsByPartNumber, tags] = await Promise.all([
      getTagsForParts(mergedParts.map((part) => part.part_number)),
      listTags(),
    ]);
    const parts = mergedParts.map((part) => ({
      ...part,
      tags: tagsByPartNumber[part.part_number] ?? [],
    }));

    const localResult = results.find((result) => result.sourceId === "local");

    const sourceStatuses: SourceStatus[] = [
      {
        id: "local",
        label: "Local catalog",
        configured: true,
        removable: false,
        ok: localResult?.ok,
        partCount: localResult?.parts.length,
        note: localResult && !localResult.ok ? localResult.note : undefined,
        debugDetail: localResult && !localResult.ok ? localResult.debugDetail : undefined,
      },
      ...sharePointEntries.map((entry): SourceStatus => {
        const result = results.find((item) => item.sourceId === entry.id);
        return {
          id: entry.id,
          label: entry.label,
          detail: entry.detail,
          configured: true,
          removable: entry.removable,
          ok: result?.ok,
          partCount: result?.parts.length,
          note: result
            ? result.ok
              ? undefined
              : result.note
            : "Signed-in session has no Microsoft access token yet — sign out and back in.",
          debugDetail: result && !result.ok ? result.debugDetail : undefined,
          fileUrl: result?.fileUrl,
        };
      }),
    ];

    return {
      parts,
      jobs: results.flatMap((result) => result.jobs),
      jobParts: results.flatMap((result) => result.jobParts),
      warnings,
      sourceStatuses,
      canAddSharePointSource: isSourceStoreConfigured(),
      tags,
      canManageTags: isTagStoreConfigured(),
    };
  },
);
