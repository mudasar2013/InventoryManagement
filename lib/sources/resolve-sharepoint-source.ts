import type { SharePointColumnMap, SharePointExcelConfig } from "./sharepoint-excel-source";
import { readSharePointExcelConfig } from "./sharepoint-excel-source";
import type { StoredSharePointSource } from "./sharepoint-source-store";
import { listStoredSharePointSources } from "./sharepoint-source-store";

/** Id of the single legacy SharePoint source configured via environment
 *  variables (see readSharePointExcelConfig) — fixed, unlike stored
 *  sources' randomly generated ids, since there's only ever one of it.
 *  Shared between lib/getInventory.ts (which builds this id's
 *  SourceStatus/InventorySource for display) and this module (which
 *  resolves it back to a config for the write API routes), so the two
 *  can't drift apart. */
export const ENV_SHAREPOINT_SOURCE_ID = "sharepoint-env";

/**
 * Builds a source's own column map from its stored per-source header
 * fields — see StoredSharePointSource.partNumberColumn etc. Returns
 * undefined (letting createSharePointExcelSource fall back to the
 * legacy default COLUMN_MAP) for a source added before per-source
 * column mapping existed, i.e. it has neither field set.
 *
 * Lives here (not in lib/getInventory.ts, which used to define this
 * locally) so this module can use it too without importing
 * getInventory.ts, which would create sources/ -> getInventory ->
 * sources/ import cycle.
 */
export function columnMapForStoredSource(
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
    category: entry.categoryColumn,
  };
}

export interface ResolvedSharePointSource {
  config: SharePointExcelConfig;
  label: string;
}

/**
 * Finds the SharePoint config for a given source id — either the
 * legacy env-configured source (ENV_SHAREPOINT_SOURCE_ID) or one added
 * from the "Data sources" page. Used by the parts update/add API
 * routes (see app/api/parts/route.ts) to turn the sourceId a technician
 * picked back into something updatePartInWorkbook/addPartToWorkbook can
 * use. Returns null when the id doesn't match a SharePoint source at
 * all (e.g. "local", or a since-removed source) — callers should treat
 * that as "nothing to write to", not attempt a Graph call.
 */
export async function resolveSharePointConfigById(
  sourceId: string,
): Promise<ResolvedSharePointSource | null> {
  if (sourceId === ENV_SHAREPOINT_SOURCE_ID) {
    const envConfig = readSharePointExcelConfig();
    return envConfig ? { config: envConfig, label: "SharePoint workbook (env)" } : null;
  }

  const stored = await listStoredSharePointSources();
  const entry = stored.find((source) => source.id === sourceId);
  if (!entry) {
    return null;
  }
  return {
    config: {
      siteHostname: entry.siteHostname,
      sitePath: entry.sitePath,
      filePath: entry.filePath,
      tableName: entry.tableName,
      columnMap: columnMapForStoredSource(entry),
    },
    label: entry.label,
  };
}
