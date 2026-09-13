import { cache } from "react";
import { describeError, describeErrorDetail } from "./describe-error";
import { localSource } from "./sources/local-source";
import { mergeParts } from "./sources/merge";
import {
  createSharePointExcelSource,
  readSharePointExcelConfig,
} from "./sources/sharepoint-excel-source";
import type { SharePointColumnMap } from "./sources/sharepoint-excel-source";
import {
  isSourceStoreConfigured,
  listStoredSharePointSources,
} from "./sources/sharepoint-source-store";
import type { StoredSharePointSource } from "./sources/sharepoint-source-store";
import type { InventorySource } from "./sources/types";
import type { Job, JobPart, Part } from "./types";

/**
 * Builds this source's own column map from its stored per-source
 * header fields — see StoredSharePointSource.partNumberColumn etc.
 * Returns undefined (letting createSharePointExcelSource fall back to
 * the legacy default COLUMN_MAP) for a source added before per-source
 * column mapping existed, i.e. it has neither field set.
 */
function columnMapForStoredSource(
  entry: StoredSharePointSource,
): SharePointColumnMap | undefined {
  if (!entry.partNumberColumn && !entry.quantityColumn) {
    return undefined;
  }
  return {
    part_number: entry.partNumberColumn || "PartNumber",
    quantity_on_hand: entry.quantityColumn || "QtyOnHand",
    description: entry.descriptionColumn,
    bin_location: entry.binLocationColumn,
    id: entry.idColumn,
  };
}

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
}

interface SharePointEntry {
  id: string;
  label: string;
  detail: string;
  removable: boolean;
}

async function buildSources(accessToken: string | undefined): Promise<{
  sources: InventorySource[];
  priority: string[];
  sharePointEntries: SharePointEntry[];
}> {
  const sources: InventorySource[] = [localSource];
  const priority: string[] = [];
  const sharePointEntries: SharePointEntry[] = [];

  // The original single SharePoint source, still configured via
  // environment variables. Kept working for whoever set this up before
  // per-source storage existed — it just can't be edited or removed from
  // the page, only by changing env vars and redeploying.
  const envConfig = readSharePointExcelConfig();
  if (envConfig) {
    const id = "sharepoint-env";
    sharePointEntries.push({
      id,
      label: "SharePoint workbook (env)",
      detail: `${envConfig.siteHostname}${envConfig.sitePath} · ${envConfig.filePath} (table/sheet: ${envConfig.tableName})`,
      removable: false,
    });
    priority.push(id);
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
    priority.push(entry.id);
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

  // Every SharePoint entry outranks the bundled demo catalog when they
  // report the same part_number — see lib/sources/merge.ts. Among
  // multiple SharePoint sources, earlier-added ones win ties, since
  // that's the order they appear in `priority`.
  priority.push("local");

  return { sources, priority, sharePointEntries };
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
    const { sources, priority, sharePointEntries } = await buildSources(accessToken);
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
          };
        }
      }),
    );

    const parts = mergeParts(
      results.map(({ sourceId, parts }) => ({ sourceId, parts })),
      priority,
    );

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
    };
  },
);
