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
  /** Path to the workbook, relative to the site's default document
   *  library's own root — e.g. just "Inventory.xlsx" for a file sitting
   *  at the top of that library, or "Team Channel/Inventory.xlsx" for
   *  one in a subfolder. Deliberately NOT including the library's own
   *  folder name ("Shared Documents", shown in the SharePoint UI as
   *  "Documents") — that's needed in a *browser* URL (which is relative
   *  to the site), but the Microsoft Graph `/drive/root:/` addressing
   *  used here is relative to the drive, and the drive already *is*
   *  that library. Pasting a browser URL's path in verbatim is the most
   *  common way to get this wrong, so createSharePointExcelSource also
   *  tries stripping a leading "Shared Documents/"/"Documents/" segment
   *  if the path as given 404s — see resolveFilePath below. */
  filePath: string;
  /** Either an Excel Table's name (Insert > Table, then named in the
   *  Table Design tab) or, when the workbook has no such Table, a
   *  worksheet's tab name — createSharePointExcelSource tries the name
   *  as a Table first and falls back to reading that worksheet's whole
   *  used range if no Table by that name exists. A named Table is more
   *  robust to inserted/removed rows, so it's worth converting a plain
   *  sheet to one (select the data, Insert > Table) if you can — but
   *  plenty of workbooks are just a sheet with headers in row 1, and
   *  this reads those too rather than requiring the conversion. */
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
 * Runs one Graph call and, on failure, rethrows wrapped in a plain Error
 * naming what was being attempted — with the original error attached as
 * `.cause` so describeErrorDetail() still shows the full Graph SDK
 * detail (status code, error code, body). A bare "itemNotFound" is
 * ambiguous across several unrelated things this source looks up (the
 * site, the table, the worksheet) — this turns it into "site lookup
 * failed" vs. "reading the table's rows failed", which is the
 * difference between checking the hostname/site path and checking the
 * file path/table name.
 */
async function runStep<T>(description: string, fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (error) {
    throw new Error(description, { cause: error });
  }
}

/** True for the specific Graph error Microsoft returns when the thing
 *  you asked for by name (a Table, a worksheet, ...) doesn't exist —
 *  as opposed to a permissions error, a network failure, etc., which
 *  should propagate rather than be treated as "try the next thing". */
export function isItemNotFoundError(error: unknown): boolean {
  return (
    error !== null &&
    typeof error === "object" &&
    (error as { code?: string | null }).code === "itemNotFound"
  );
}

/** The default document library's folder name(s) SharePoint shows in a
 *  browser URL — lowercased for a case-insensitive match. */
const DEFAULT_LIBRARY_FOLDER_NAMES = ["shared documents", "documents"];

/** Strips a leading "Shared Documents/" or "Documents/" segment from a
 *  file path, if present — returns null when there's nothing to strip
 *  (so the caller can tell "no second candidate" from "stripped to an
 *  empty string"). See the filePath doc comment on SharePointExcelConfig
 *  for why this mix-up happens and why it's worth trying both. */
function withoutDefaultLibraryPrefix(filePath: string): string | null {
  const firstSlash = filePath.indexOf("/");
  if (firstSlash === -1) {
    return null;
  }
  const firstSegment = filePath.slice(0, firstSlash).trim().toLowerCase();
  if (!DEFAULT_LIBRARY_FOLDER_NAMES.includes(firstSegment)) {
    return null;
  }
  const rest = filePath.slice(firstSlash + 1);
  return rest.length > 0 ? rest : null;
}

/**
 * Confirms which of one or two candidate paths the workbook actually
 * lives at — the path as configured, and (when it starts with a default
 * library folder name) that same path with the folder name stripped —
 * and returns whichever one resolves. A wrong path shows up as
 * "itemNotFound" here just like a wrong table/worksheet name would
 * further in, so this runs first and gets its own clear error rather
 * than leaving the ambiguity to whatever fails next.
 */
export async function resolveFilePath(
  client: Client,
  siteId: string,
  filePath: string,
): Promise<string> {
  const stripped = withoutDefaultLibraryPrefix(filePath);
  const candidates = stripped ? [filePath, stripped] : [filePath];

  let lastError: unknown;
  for (const candidate of candidates) {
    try {
      await client.api(`/sites/${siteId}/drive/root:/${encodeURI(candidate)}`).get();
      return candidate;
    } catch (error) {
      if (!isItemNotFoundError(error)) {
        throw new Error(`Looking up file "${candidate}" failed`, { cause: error });
      }
      lastError = error;
    }
  }

  throw new Error(
    `File not found at "${filePath}"` +
      (stripped ? ` (also tried "${stripped}")` : "") +
      ` — check "File path" on the Data sources page matches where the file actually is, ` +
      `relative to the site's default document library's own root (its own folder name, ` +
      `"Shared Documents" / "Documents", is not part of that path)`,
    { cause: lastError },
  );
}

/**
 * Reads a workbook's data as headers + rows, trying an Excel Table
 * first and falling back to a plain worksheet's used range if no Table
 * by that name exists. Most shop inventory workbooks are just a sheet
 * with headers in row 1 rather than a formally-named Table, and
 * requiring the conversion before this source could read anything was
 * a real usability wall — this tries the friendlier, more robust
 * option first and quietly falls back rather than making the
 * conversion a precondition.
 */
export async function readWorkbookData(
  client: Client,
  fileBase: string,
  filePath: string,
  tableOrWorksheetName: string,
): Promise<{ headers: string[]; rows: unknown[][] }> {
  const tableBase = `${fileBase}/tables/${encodeURIComponent(tableOrWorksheetName)}`;

  let tableExists: boolean;
  try {
    await client.api(tableBase).get();
    tableExists = true;
  } catch (error) {
    if (!isItemNotFoundError(error)) {
      throw new Error(`Looking up Excel Table "${tableOrWorksheetName}" failed`, {
        cause: error,
      });
    }
    tableExists = false;
  }

  if (tableExists) {
    const [headerRange, rowsResponse] = await Promise.all([
      runStep(`Reading Table "${tableOrWorksheetName}"'s header row failed`, () =>
        client.api(`${tableBase}/headerRowRange`).get(),
      ),
      runStep(`Reading Table "${tableOrWorksheetName}"'s rows failed`, () =>
        client.api(`${tableBase}/rows`).get(),
      ),
    ]);
    return {
      headers: (headerRange.values?.[0] ?? []).map((value: unknown) => String(value)),
      rows: (rowsResponse.value ?? []).map((row: { values: unknown[][] }) => row.values[0]),
    };
  }

  // No Excel Table by that name — try it as a worksheet tab name
  // instead, reading every used cell directly (row 1 is assumed to be
  // headers, same as the Table path). The file itself is already
  // confirmed to exist by this point (resolveFilePath runs before this
  // is called), so a 404 here means "Table name" doesn't match either
  // an Excel Table or a worksheet tab name.
  const usedRange = await runStep(
    `Reading worksheet "${tableOrWorksheetName}" in "${filePath}" failed — check that ` +
      `"Table name" on the Data sources page matches either an actual Excel Table name ` +
      `(Insert > Table, named in the Table Design tab) or a worksheet tab name at the ` +
      `bottom of Excel`,
    () =>
      client
        .api(`${fileBase}/worksheets/${encodeURIComponent(tableOrWorksheetName)}/usedRange`)
        .get(),
  );

  const values: unknown[][] = usedRange.values ?? [];
  return {
    headers: (values[0] ?? []).map((value: unknown) => String(value)),
    rows: values.slice(1),
  };
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
      const site = await runStep(
        `Site lookup failed for "${config.siteHostname}${config.sitePath}" — check the ` +
          `"Site hostname" and "Site path" fields on the Data sources page`,
        () => client.api(`/sites/${config.siteHostname}:${config.sitePath}`).get(),
      );

      const resolvedFilePath = await resolveFilePath(client, site.id, config.filePath);
      const fileBase = `/sites/${site.id}/drive/root:/${encodeURI(resolvedFilePath)}:/workbook`;
      const { headers, rows } = await readWorkbookData(
        client,
        fileBase,
        resolvedFilePath,
        config.tableName,
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
