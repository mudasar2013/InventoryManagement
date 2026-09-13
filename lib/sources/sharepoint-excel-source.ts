import { Client } from "@microsoft/microsoft-graph-client";
import type { ExtraFields, RawPart } from "../types";
import type { InventorySource } from "./types";

/**
 * Which row-1 header text to read each RawPart field from. Every
 * workbook is free to name (and order) its columns however it likes —
 * this is what lets each SharePoint source have its own layout while
 * still feeding the same shape into the merged catalog (see
 * lib/sources/merge.ts, which unions/dedupes every source's RawParts by
 * part_number into one master list). `part_number` and
 * `quantity_on_hand` are required — mapTableRowsToRawParts throws a
 * clear error naming what it looked for if either header isn't found,
 * rather than silently reading garbage. The rest are optional; a
 * missing optional column just comes through blank (or, for `id`,
 * synthesized from part_number).
 */
export interface SharePointColumnMap {
  part_number: string;
  description?: string;
  bin_location?: string;
  quantity_on_hand: string;
  id?: string;
  /** Row-1 header this source uses for category, when it has one — see
   *  RawPart.category. Every other column not named by one of this
   *  map's fields is captured automatically as an extra field (see
   *  ExtraFields), so category gets its own entry only because it's
   *  common enough to deserve a first-class filter on the Parts page. */
  category?: string;
}

/**
 * Fallback column map used when a source doesn't specify its own —
 * the legacy env-configured source, and any source added before
 * per-source column mapping existed. New sources added from the "Data
 * sources" page always specify their own (see SharePointExcelConfig.
 * columnMap below), since assuming every workbook shares one fixed set
 * of header names is exactly the assumption that broke on a second,
 * differently-laid-out sheet.
 */
export const COLUMN_MAP: SharePointColumnMap = {
  part_number: "PartNumber",
  description: "Description",
  bin_location: "BinLocation",
  quantity_on_hand: "QtyOnHand",
  id: undefined,
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
  /** This source's own header-name mapping — see SharePointColumnMap.
   *  Falls back to COLUMN_MAP (the legacy hardcoded default) when
   *  omitted, for the env-configured source and any source added
   *  before this field existed. */
  columnMap?: SharePointColumnMap;
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

/** Finds a header's column index by exact text match (after trimming),
 *  or -1 when `header` is undefined (this field isn't mapped to any
 *  column) or not found. Shared by the read path (mapTableRowsToRawParts)
 *  and the write path (updatePartInWorkbook / addPartToWorkbook) so both
 *  agree on what "this column isn't in the sheet" means. */
function headerIndex(headers: string[], header: string | undefined): number {
  return header === undefined ? -1 : headers.findIndex((h) => String(h).trim() === header);
}

/** Spellings a shop-floor sheet uses for a yes/no cell — checked
 *  case-insensitively, trimmed. Used to both detect whether a whole
 *  column is boolean-shaped (see columnLooksBoolean) and to read one
 *  cell's boolean value once a column has been classified as such. */
const BOOLEAN_CELL_VALUES = new Set(["YES", "NO", "Y", "N", "TRUE", "FALSE"]);
const TRUE_CELL_VALUES = new Set(["YES", "Y", "TRUE"]);

/** A column "looks boolean" when every non-blank cell it has is some
 *  spelling of yes/no — e.g. "Ebay Ready (Yes/No)", "In Inventory",
 *  "CORE RETURNED", "Sold". A column that's entirely blank in the rows
 *  seen doesn't count (nothing to infer from; safer to show it as
 *  text than to guess). Checked once per column across every row,
 *  rather than per cell, so one part's blank cell in an otherwise
 *  yes/no column still renders as a checkbox rather than flickering to
 *  a text field depending on which row happens to be blank. */
function columnLooksBoolean(rows: unknown[][], columnIndex: number): boolean {
  let sawValue = false;
  for (const row of rows) {
    const trimmed = String(row[columnIndex] ?? "").trim();
    if (trimmed === "") continue;
    sawValue = true;
    if (!BOOLEAN_CELL_VALUES.has(trimmed.toUpperCase())) {
      return false;
    }
  }
  return sawValue;
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
 *  should propagate rather than be treated as "try the next thing".
 *  Case-insensitive: the drive/file endpoints used by resolveFilePath
 *  return "itemNotFound", but the workbook Tables endpoint has been
 *  observed returning "ItemNotFound" (capital I) for the exact same
 *  situation — an inconsistency in Graph's own Excel API, not
 *  something this code can fix, only work around. */
export function isItemNotFoundError(error: unknown): boolean {
  if (error === null || typeof error !== "object") {
    return false;
  }
  const code = (error as { code?: string | null }).code;
  return typeof code === "string" && code.toLowerCase() === "itemnotfound";
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

/** The file's resolved drive-relative path, plus its real SharePoint
 *  browser URL (the DriveItem's own `webUrl`, straight from Graph —
 *  more reliable than hand-building one from siteHostname/sitePath/
 *  filePath, since those don't by themselves tell you which library
 *  folder name convention actually applies). `webUrl` is left
 *  undefined on the rare response that omits it rather than failing
 *  the whole lookup over a "nice to have" field. */
export interface ResolvedFile {
  path: string;
  webUrl?: string;
}

/**
 * Confirms which of one or two candidate paths the workbook actually
 * lives at — the path as configured, and (when it starts with a default
 * library folder name) that same path with the folder name stripped —
 * and returns whichever one resolves, together with that file's real
 * webUrl for "open this file" links (see createSharePointExcelSource's
 * getFileUrl). A wrong path shows up as "itemNotFound" here just like a
 * wrong table/worksheet name would further in, so this runs first and
 * gets its own clear error rather than leaving the ambiguity to
 * whatever fails next.
 */
export async function resolveFilePath(
  client: Client,
  siteId: string,
  filePath: string,
): Promise<ResolvedFile> {
  const stripped = withoutDefaultLibraryPrefix(filePath);
  const candidates = stripped ? [filePath, stripped] : [filePath];

  let lastError: unknown;
  for (const candidate of candidates) {
    try {
      const item = await client
        .api(`/sites/${siteId}/drive/root:/${encodeURI(candidate)}`)
        .get();
      return {
        path: candidate,
        webUrl: typeof item?.webUrl === "string" ? item.webUrl : undefined,
      };
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

/** Checks whether an Excel Table by this name exists at `tableBase`,
 *  without throwing for the "it just doesn't exist" case — used by both
 *  the read path (readWorkbookData) and the write path
 *  (updatePartInWorkbook / addPartToWorkbook) to decide whether to use
 *  the Table APIs or fall back to a plain worksheet. */
async function probeTableExists(
  client: Client,
  tableBase: string,
  tableOrWorksheetName: string,
): Promise<boolean> {
  try {
    await client.api(tableBase).get();
    return true;
  } catch (error) {
    if (!isItemNotFoundError(error)) {
      throw new Error(`Looking up Excel Table "${tableOrWorksheetName}" failed`, {
        cause: error,
      });
    }
    return false;
  }
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
  const tableExists = await probeTableExists(client, tableBase, tableOrWorksheetName);

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
  columnMap: SharePointColumnMap = COLUMN_MAP,
): RawPart[] {
  const partNumberIdx = headerIndex(headers, columnMap.part_number);
  const descriptionIdx = headerIndex(headers, columnMap.description);
  const binLocationIdx = headerIndex(headers, columnMap.bin_location);
  const quantityIdx = headerIndex(headers, columnMap.quantity_on_hand);
  const idIdx = headerIndex(headers, columnMap.id);
  const categoryIdx = headerIndex(headers, columnMap.category);

  if (partNumberIdx === -1 || quantityIdx === -1) {
    throw new Error(
      `SharePoint workbook is missing an expected column. Looked for ` +
        `"${columnMap.part_number}" and "${columnMap.quantity_on_hand}" ` +
        `among headers: ${headers.join(", ") || "(none found)"}. Set "Part number ` +
        `column" and "Quantity column" on the Data sources page to match the real ` +
        `headers for this source.`,
    );
  }

  // Every header not already claimed by a named field (part number,
  // description, bin location, quantity, id, category) becomes an
  // extra field — see RawPart.extraFields. Classifying which of those
  // are boolean-shaped is done once per column, across every row, not
  // per cell (see columnLooksBoolean).
  const mappedIndices = new Set(
    [partNumberIdx, descriptionIdx, binLocationIdx, quantityIdx, idIdx, categoryIdx].filter(
      (idx) => idx !== -1,
    ),
  );
  const extraIndices = headers
    .map((header, idx) => ({ header: String(header ?? "").trim(), idx }))
    .filter(({ header, idx }) => header.length > 0 && !mappedIndices.has(idx));
  const booleanIndices = new Set(
    extraIndices
      .map(({ idx }) => idx)
      .filter((idx) => columnLooksBoolean(rows, idx)),
  );

  return rows
    .filter((row) => String(row[partNumberIdx] ?? "").trim().length > 0)
    .map((row) => {
      const partNumber = String(row[partNumberIdx]).trim();
      const quantity = Number(row[quantityIdx]);

      const extraFields: ExtraFields = {};
      for (const { header, idx } of extraIndices) {
        const cell = row[idx];
        if (booleanIndices.has(idx)) {
          const normalized = String(cell ?? "").trim().toUpperCase();
          extraFields[header] = { kind: "boolean", value: TRUE_CELL_VALUES.has(normalized) };
        } else {
          extraFields[header] = { kind: "text", value: String(cell ?? "") };
        }
      }

      return {
        id: idIdx !== -1 ? String(row[idIdx]) : `sharepoint-${slugify(partNumber)}`,
        part_number: partNumber,
        description: descriptionIdx !== -1 ? String(row[descriptionIdx] ?? "") : "",
        bin_location: binLocationIdx !== -1 ? String(row[binLocationIdx] ?? "") : "",
        quantity_on_hand: Number.isFinite(quantity) ? quantity : 0,
        category: categoryIdx !== -1 ? String(row[categoryIdx] ?? "") : undefined,
        extraFields,
      };
    });
}

/**
 * Resolves a config down to a ready-to-use Graph client and the
 * `.../workbook` API base for its file — the site lookup + file-path
 * resolution (see resolveFilePath) that both reading and writing need
 * before they can do anything else. Shared so the write path
 * (updatePartInWorkbook / addPartToWorkbook, called from the "add/edit
 * a part" API routes) doesn't re-implement the same site/path
 * resolution createSharePointExcelSource's fetchParts already does.
 */
export async function connectToWorkbook(
  accessToken: string,
  config: SharePointExcelConfig,
): Promise<{ client: Client; fileBase: string; filePath: string; fileUrl?: string }> {
  const client = Client.init({
    authProvider: (done) => done(null, accessToken),
  });

  const site = await runStep(
    `Site lookup failed for "${config.siteHostname}${config.sitePath}" — either the ` +
      `"Site hostname"/"Site path" fields on the Data sources page are wrong, or (with ` +
      `Sites.Selected permissions) this app hasn't been granted access to this specific ` +
      `site yet — see the README's "Sign-in setup" section for the ` +
      `Grant-PnPAzureADAppSitePermission command`,
    () => client.api(`/sites/${config.siteHostname}:${config.sitePath}`).get(),
  );

  const resolved = await resolveFilePath(client, site.id, config.filePath);
  return {
    client,
    fileBase: `/sites/${site.id}/drive/root:/${encodeURI(resolved.path)}:/workbook`,
    filePath: resolved.path,
    fileUrl: resolved.webUrl,
  };
}

/** The part fields a technician can edit or supply for a new part —
 *  everything RawPart carries except `id`, which is either read from a
 *  sheet's own id column or synthesized from part_number, never
 *  user-entered directly. */
export type PartFields = Omit<RawPart, "id">;

/** Converts a 0-based column index to spreadsheet column letters
 *  (0 -> "A", 25 -> "Z", 26 -> "AA", ...) — used to address a single
 *  cell by A1 notation when writing to a plain worksheet, since a
 *  source's mapped columns (part number, quantity, ...) are frequently
 *  scattered rather than contiguous and can't be written as one range. */
function columnIndexToLetters(index: number): string {
  let n = index;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return letters;
}

/** Builds the (header, value) pairs to write for one part's fields —
 *  the named fields (part number, description, ...) by their mapped
 *  column, plus every entry in `fields.extraFields` by that entry's own
 *  header text (extra fields aren't looked up through columnMap at
 *  all: the header they were read from, see mapTableRowsToRawParts, is
 *  exactly the header they're written back to). Shared by both the
 *  Table row path (applyFieldsToRow) and the plain-worksheet path
 *  (writeFieldsToWorksheetRow) so the two can't drift on what counts
 *  as "this field has a column to write to". A boolean extra field is
 *  written back as "YES"/"NO", matching how these sheets already
 *  spell it (see BOOLEAN_CELL_VALUES) rather than a literal true/false
 *  a shop-floor sheet has never seen. */
function fieldAssignments(
  columnMap: SharePointColumnMap,
  fields: PartFields,
): [string | undefined, unknown][] {
  const assignments: [string | undefined, unknown][] = [
    [columnMap.part_number, fields.part_number],
    [columnMap.description, fields.description],
    [columnMap.bin_location, fields.bin_location],
    [columnMap.quantity_on_hand, fields.quantity_on_hand],
  ];
  if (columnMap.category !== undefined) {
    assignments.push([columnMap.category, fields.category ?? ""]);
  }
  for (const [header, field] of Object.entries(fields.extraFields ?? {})) {
    assignments.push([
      header,
      field.kind === "boolean" ? (field.value ? "YES" : "NO") : field.value,
    ]);
  }
  return assignments;
}

/** Overwrites the mapped cells of an existing row array in place —
 *  shared by the Table update and Table add-row paths, both of which
 *  work with a full-width row array matching the table's own header
 *  order. Fields with no configured column are left untouched (nothing
 *  in the sheet to put them in). */
function applyFieldsToRow(
  row: unknown[],
  headers: string[],
  columnMap: SharePointColumnMap,
  fields: PartFields,
): void {
  for (const [header, value] of fieldAssignments(columnMap, fields)) {
    const idx = headerIndex(headers, header);
    if (idx !== -1) {
      row[idx] = value;
    }
  }
}

/** Writes each mapped field to its own single-cell range in a plain
 *  worksheet (no Table) — one Graph call per configured field, rather
 *  than one range covering the whole row, because a source's mapped
 *  columns are frequently scattered across the sheet rather than
 *  contiguous (see the "Test Data" source this was built for: PartName,
 *  PartNumber and Quantity aren't next to each other). Fields with no
 *  configured column are skipped — nowhere in the sheet to write them. */
async function writeFieldsToWorksheetRow(
  client: Client,
  worksheetBase: string,
  worksheetName: string,
  headers: string[],
  startColumn0: number,
  rowIndex0: number,
  columnMap: SharePointColumnMap,
  fields: PartFields,
): Promise<void> {
  for (const [header, value] of fieldAssignments(columnMap, fields)) {
    const colIdx = headerIndex(headers, header);
    if (colIdx === -1) {
      continue;
    }
    const address = `${columnIndexToLetters(startColumn0 + colIdx)}${rowIndex0 + 1}`;
    await runStep(`Writing "${header}" to worksheet "${worksheetName}" failed`, () =>
      client
        .api(`${worksheetBase}/range(address='${address}')`)
        .patch({ values: [[value]] }),
    );
  }
}

/**
 * Updates the row for `existingPartNumber` in place — an Excel Table
 * row via a single PATCH covering the whole (table-relative, so always
 * contiguous) row, or a plain worksheet's scattered columns one cell at
 * a time (see writeFieldsToWorksheetRow). Throws a clear error if no
 * row with that part number is found, or if the part-number column
 * itself isn't configured/found (there'd be nothing to match against).
 */
export async function updatePartInWorkbook(
  client: Client,
  fileBase: string,
  tableOrWorksheetName: string,
  columnMap: SharePointColumnMap,
  existingPartNumber: string,
  fields: PartFields,
): Promise<void> {
  const tableBase = `${fileBase}/tables/${encodeURIComponent(tableOrWorksheetName)}`;
  const tableExists = await probeTableExists(client, tableBase, tableOrWorksheetName);

  if (tableExists) {
    const headerRange = await runStep(
      `Reading Table "${tableOrWorksheetName}"'s header row failed`,
      () => client.api(`${tableBase}/headerRowRange`).get(),
    );
    const headers: string[] = (headerRange.values?.[0] ?? []).map((v: unknown) => String(v));
    const partNumberIdx = headerIndex(headers, columnMap.part_number);
    if (partNumberIdx === -1) {
      throw new Error(
        `Can't update this part: no "${columnMap.part_number}" column found in Table ` +
          `"${tableOrWorksheetName}".`,
      );
    }

    const rowsResponse = await runStep(
      `Reading Table "${tableOrWorksheetName}"'s rows failed`,
      () => client.api(`${tableBase}/rows`).get(),
    );
    const tableRows: { index: number; values: unknown[][] }[] = rowsResponse.value ?? [];
    const match = tableRows.find(
      (row) => String(row.values[0]?.[partNumberIdx] ?? "").trim() === existingPartNumber,
    );
    if (!match) {
      throw new Error(
        `Can't update this part: no row with part number "${existingPartNumber}" found ` +
          `in Table "${tableOrWorksheetName}".`,
      );
    }

    const newValues = [...match.values[0]];
    applyFieldsToRow(newValues, headers, columnMap, fields);

    await runStep(
      `Writing the updated row for "${existingPartNumber}" failed`,
      () => client.api(`${tableBase}/rows/itemAt(index=${match.index})`).patch({
        values: [newValues],
      }),
    );
    return;
  }

  const worksheetBase = `${fileBase}/worksheets/${encodeURIComponent(tableOrWorksheetName)}`;
  const usedRange = await runStep(
    `Reading worksheet "${tableOrWorksheetName}" failed`,
    () => client.api(`${worksheetBase}/usedRange`).get(),
  );
  const values: unknown[][] = usedRange.values ?? [];
  const headers = (values[0] ?? []).map((v: unknown) => String(v));
  const partNumberIdx = headerIndex(headers, columnMap.part_number);
  if (partNumberIdx === -1) {
    throw new Error(
      `Can't update this part: no "${columnMap.part_number}" column found in worksheet ` +
        `"${tableOrWorksheetName}".`,
    );
  }

  const dataRows = values.slice(1);
  const matchOffset = dataRows.findIndex(
    (row) => String(row[partNumberIdx] ?? "").trim() === existingPartNumber,
  );
  if (matchOffset === -1) {
    throw new Error(
      `Can't update this part: no row with part number "${existingPartNumber}" found in ` +
        `worksheet "${tableOrWorksheetName}".`,
    );
  }

  const startColumn0: number = usedRange.columnIndex ?? 0;
  const rowIndex0: number = (usedRange.rowIndex ?? 0) + 1 + matchOffset; // +1 skips the header row

  await writeFieldsToWorksheetRow(
    client,
    worksheetBase,
    tableOrWorksheetName,
    headers,
    startColumn0,
    rowIndex0,
    columnMap,
    fields,
  );
}

/**
 * Adds a brand-new row to the workbook — appended to the Excel Table
 * (via the Table "add row" API, so it becomes part of the Table just
 * like a row typed in by hand) or written just past the current used
 * range for a plain worksheet.
 */
export async function addPartToWorkbook(
  client: Client,
  fileBase: string,
  tableOrWorksheetName: string,
  columnMap: SharePointColumnMap,
  fields: PartFields,
): Promise<void> {
  const tableBase = `${fileBase}/tables/${encodeURIComponent(tableOrWorksheetName)}`;
  const tableExists = await probeTableExists(client, tableBase, tableOrWorksheetName);

  if (tableExists) {
    const headerRange = await runStep(
      `Reading Table "${tableOrWorksheetName}"'s header row failed`,
      () => client.api(`${tableBase}/headerRowRange`).get(),
    );
    const headers: string[] = (headerRange.values?.[0] ?? []).map((v: unknown) => String(v));
    const newRow: unknown[] = headers.map(() => "");
    applyFieldsToRow(newRow, headers, columnMap, fields);

    await runStep(
      `Adding a new row to Table "${tableOrWorksheetName}" failed`,
      () => client.api(`${tableBase}/rows/add`).post({ values: [newRow] }),
    );
    return;
  }

  const worksheetBase = `${fileBase}/worksheets/${encodeURIComponent(tableOrWorksheetName)}`;
  const usedRange = await runStep(
    `Reading worksheet "${tableOrWorksheetName}" failed`,
    () => client.api(`${worksheetBase}/usedRange`).get(),
  );
  const values: unknown[][] = usedRange.values ?? [];
  const headers = (values[0] ?? []).map((v: unknown) => String(v));
  const startColumn0: number = usedRange.columnIndex ?? 0;
  const newRowIndex0: number = (usedRange.rowIndex ?? 0) + values.length;

  await writeFieldsToWorksheetRow(
    client,
    worksheetBase,
    tableOrWorksheetName,
    headers,
    startColumn0,
    newRowIndex0,
    columnMap,
    fields,
  );
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
  // Populated by fetchParts() once the site + file lookups succeed —
  // which happens before the table/worksheet is even read, so this is
  // available for "open this file" links even when the fetch goes on
  // to fail on a bad table/worksheet name. Stays undefined until
  // fetchParts() has actually run at least once this request.
  let resolvedFileUrl: string | undefined;

  return {
    id: identity?.id ?? "sharepoint",
    label: identity?.label ?? `SharePoint workbook (${config.filePath})`,
    async fetchParts(): Promise<RawPart[]> {
      const { client, fileBase, filePath, fileUrl } = await connectToWorkbook(
        accessToken,
        config,
      );
      resolvedFileUrl = fileUrl;
      const { headers, rows } = await readWorkbookData(client, fileBase, filePath, config.tableName);

      return mapTableRowsToRawParts(headers, rows, config.columnMap ?? COLUMN_MAP);
    },
    getFileUrl() {
      return resolvedFileUrl;
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
