import { Client } from "@microsoft/microsoft-graph-client";
import type { RawPart } from "../types";
import type { InventorySource } from "./types";

/**
 * ---------------------------------------------------------------------
 * EDIT THIS to match the real workbook's column headers once you can
 * see it. `part_number` and `quantity_on_hand` are required — the
 * source throws a clear error naming what it looked for if either is
 * missing, rather than silently reading garbage. The rest are optional;
 * a missing optional column just comes through blank.
 * ---------------------------------------------------------------------
 */
export const COLUMN_MAP = {
  part_number: "PartNumber",
  description: "Description",
  bin_location: "BinLocation",
  quantity_on_hand: "QtyOnHand",
  // Set this to a header name only if the sheet has its own stable id
  // column. Leave it undefined (the default) to synthesize one from
  // part_number instead — most inventory spreadsheets don't have one.
  id: undefined as string | undefined,
};

export interface SharePointExcelConfig {
  /** e.g. "contoso.sharepoint.com" */
  siteHostname: string;
  /** Server-relative site path, e.g. "/sites/ServiceOps" */
  sitePath: string;
  /** Path to the workbook within the site's default document library,
   *  e.g. "Shared Documents/Inventory.xlsx" */
  filePath: string;
  /** The Excel Table's name (Insert > Table, then name it in the Table
   *  Design tab) — not the worksheet name. A named Table is what makes
   *  this robust to inserted rows/columns; a plain cell range is not. */
  tableName: string;
}

/**
 * Reads SHAREPOINT_SITE_HOSTNAME / SHAREPOINT_SITE_PATH /
 * SHAREPOINT_FILE_PATH / SHAREPOINT_TABLE_NAME from the environment.
 * Returns null (source disabled, local catalog only) until all four are
 * set — see .env.example.
 */
export function readSharePointExcelConfig(): SharePointExcelConfig | null {
  const siteHostname = process.env.SHAREPOINT_SITE_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;
  const filePath = process.env.SHAREPOINT_FILE_PATH;
  const tableName = process.env.SHAREPOINT_TABLE_NAME;

  if (!siteHostname || !sitePath || !filePath || !tableName) {
    return null;
  }
  return { siteHostname, sitePath, filePath, tableName };
}

function slugify(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
}

/**
 * Turns an Excel Table's header row + row values into RawPart records.
 * Pulled out as a pure function, independent of the Graph client, so
 * it's unit-testable without a live SharePoint connection — the actual
 * network shape (headers: string[], rows: unknown[][]) is exactly what
 * the Graph workbook range/rows APIs return.
 */
export function mapTableRowsToRawParts(
  headers: string[],
  rows: unknown[][],
  columnMap: typeof COLUMN_MAP = COLUMN_MAP,
): RawPart[] {
  const indexOf = (header: string | undefined) =>
    header === undefined ? -1 : headers.findIndex((h) => String(h).trim() === header);

  const partNumberIdx = indexOf(columnMap.part_number);
  const descriptionIdx = indexOf(columnMap.description);
  const binLocationIdx = indexOf(columnMap.bin_location);
  const quantityIdx = indexOf(columnMap.quantity_on_hand);
  const idIdx = indexOf(columnMap.id);

  if (partNumberIdx === -1 || quantityIdx === -1) {
    throw new Error(
      `SharePoint workbook is missing an expected column. Looked for ` +
        `"${columnMap.part_number}" and "${columnMap.quantity_on_hand}" ` +
        `among headers: ${headers.join(", ") || "(none found)"}. Update ` +
        `COLUMN_MAP in lib/sources/sharepoint-excel-source.ts to match ` +
        `the real headers.`,
    );
  }

  return rows
    .filter((row) => String(row[partNumberIdx] ?? "").trim().length > 0)
    .map((row) => {
      const partNumber = String(row[partNumberIdx]).trim();
      const quantity = Number(row[quantityIdx]);
      return {
        id: idIdx !== -1 ? String(row[idIdx]) : `sharepoint-${slugify(partNumber)}`,
        part_number: partNumber,
        description: descriptionIdx !== -1 ? String(row[descriptionIdx] ?? "") : "",
        bin_location: binLocationIdx !== -1 ? String(row[binLocationIdx] ?? "") : "",
        quantity_on_hand: Number.isFinite(quantity) ? quantity : 0,
      };
    });
}

/**
 * Reads parts from an Excel workbook table via Microsoft Graph, using
 * the signed-in technician's own delegated access token (see
 * lib/auth/options.ts) — this source can only see what that person can
 * already see in SharePoint. Jobs stay local for now; nothing here
 * suggests the shop's job board also lives in this workbook.
 *
 * `identity` lets a caller give this instance its own id/label instead
 * of the "sharepoint" default — required once more than one SharePoint
 * workbook can be configured (see lib/sources/sharepoint-source-store.ts
 * and lib/getInventory.ts), since merge.ts and the priority list key
 * sources by id and two sources sharing an id would silently clobber
 * each other's provenance tracking.
 */
export function createSharePointExcelSource(
  accessToken: string,
  config: SharePointExcelConfig,
  identity?: { id?: string; label?: string },
): InventorySource {
  const client = Client.init({
    authProvider: (done) => done(null, accessToken),
  });

  return {
    id: identity?.id ?? "sharepoint",
    label: identity?.label ?? `SharePoint workbook (${config.filePath})`,
    async fetchParts(): Promise<RawPart[]> {
      const site = await client
        .api(`/sites/${config.siteHostname}:${config.sitePath}`)
        .get();

      const workbookBase = `/sites/${site.id}/drive/root:/${encodeURI(
        config.filePath,
      )}:/workbook/tables/${encodeURIComponent(config.tableName)}`;

      const [headerRange, rowsResponse] = await Promise.all([
        client.api(`${workbookBase}/headerRowRange`).get(),
        client.api(`${workbookBase}/rows`).get(),
      ]);

      const headers: string[] = (headerRange.values?.[0] ?? []).map((value: unknown) =>
        String(value),
      );
      const rows: unknown[][] = (rowsResponse.value ?? []).map(
        (row: { values: unknown[][] }) => row.values[0],
      );

      return mapTableRowsToRawParts(headers, rows);
    },
    // The workbook is inventory only — jobs and job/part links still come
    // from the local source until there's a real job system to read.
    async fetchJobs() {
      return [];
    },
    async fetchJobParts() {
      return [];
    },
  };
}
