import { Client } from "@microsoft/microsoft-graph-client";
import { CONDITION_OPTIONS } from "../types";
import type { ExtraField, ExtraFields, RawPart } from "../types";
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

/** Finds the index of the `occurrence`-th (1-based, in array order) item
 *  matching `predicate` — used by updatePartInWorkbook to land an edit on
 *  the exact physical row it came from when a part_number appears more
 *  than once in the sheet (see RawPart.partNumberOccurrence), rather than
 *  always the first such row. `occurrence` defaults to 1 everywhere it's
 *  optional, so an ordinary part_number (only one matching row) behaves
 *  exactly as it did before this existed. Returns -1 if there are fewer
 *  than `occurrence` matches. */
function nthMatchIndex<T>(items: T[], predicate: (item: T) => boolean, occurrence: number): number {
  let seen = 0;
  for (let i = 0; i < items.length; i++) {
    if (predicate(items[i])) {
      seen++;
      if (seen === occurrence) {
        return i;
      }
    }
  }
  return -1;
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

/** A header that names itself a yes/no column outright — e.g. "Ebay
 *  Ready (Yes/No)", "Ebay Listed (Yes/No)" — is boolean regardless of
 *  what this particular fetch's rows happen to contain (columnLooksBoolean
 *  needs at least one non-blank cell to go on, so a column that's
 *  entirely blank in every row seen so far would otherwise fall back to
 *  "text" even though the sheet's own header says otherwise). */
function headerNamesBoolean(header: string): boolean {
  return /\(\s*yes\s*\/\s*no\s*\)/i.test(header);
}

/** A header containing "date" (case-insensitive) — "Entry Date", "Date
 *  Rcvd", etc. These sheets store dates as raw Excel serial numbers
 *  (days since 1899-12-30); reading them as plain numbers is where "the
 *  date shows up as a number" bug came from. See excelSerialToIsoDate /
 *  isoDateToExcelSerial below for the conversion. */
function headerNamesDate(header: string): boolean {
  return /date/i.test(header);
}

/** Headers this app knows a fixed, shop-specific vocabulary for. Keyed
 *  lowercase for a case-insensitive match against the sheet's own
 *  header text. A cell value outside this list is never discarded —
 *  see ExtraField.options and the "select" kind's "Other" fallback in
 *  PartDetail — so an existing sheet with values this list doesn't (yet)
 *  cover keeps showing/editing them rather than losing data. */
const SELECT_FIELD_OPTIONS: Record<string, string[]> = {
  condition: [...CONDITION_OPTIONS],
};

function selectOptionsForHeader(header: string): string[] | undefined {
  return SELECT_FIELD_OPTIONS[header.trim().toLowerCase()];
}

/** Classifies one extra column once, from its header text and (for the
 *  boolean case, when the header itself doesn't already say so) the
 *  values actually seen in it — see the ExtraFieldKind doc comment in
 *  lib/types.ts for what each kind means. Order matters: a header this
 *  app has a fixed vocabulary for (Condition) always wins, then a
 *  header that names itself a date or a yes/no column, and only then
 *  does content-sniffing (columnLooksBoolean) get a say. */
function classifyExtraColumn(
  header: string,
  columnIndex: number,
  rows: unknown[][],
): { kind: "select"; options: string[] } | { kind: "date" } | { kind: "boolean" } | { kind: "text" } {
  const options = selectOptionsForHeader(header);
  if (options) {
    return { kind: "select", options };
  }
  if (headerNamesDate(header)) {
    return { kind: "date" };
  }
  if (headerNamesBoolean(header) || columnLooksBoolean(rows, columnIndex)) {
    return { kind: "boolean" };
  }
  return { kind: "text" };
}

/** Excel's date serial epoch: day 0 is 1899-12-30 (not 1899-12-31 —
 *  this bakes in Excel's own long-standing 1900-leap-year quirk, which
 *  only matters for dates before March 1900 and is irrelevant to any
 *  real shop-floor date). */
const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 24 * 60 * 60 * 1000;

/** Converts an Excel date serial number to a "YYYY-MM-DD" string for
 *  display/editing (an <input type="date"> wants exactly this shape). */
function excelSerialToIsoDate(serial: number): string {
  const date = new Date(EXCEL_EPOCH_UTC_MS + serial * MS_PER_DAY);
  return date.toISOString().slice(0, 10);
}

/** The inverse of excelSerialToIsoDate — what gets written back to the
 *  sheet so the cell round-trips as the same kind of value it was read
 *  as (a plain serial number), not a string Excel would have to
 *  re-parse. */
function isoDateToExcelSerial(iso: string): number {
  const [year, month, day] = iso.split("-").map(Number);
  return Math.round((Date.UTC(year, month - 1, day) - EXCEL_EPOCH_UTC_MS) / MS_PER_DAY);
}

/** Today's date as "YYYY-MM-DD", in the server's local time zone — used
 *  to auto-populate an "Entry Date"-shaped column when a brand-new part
 *  is added (see withAutoEntryDate below), so nobody has to remember to
 *  set it by hand. */
function todayIsoDate(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

/** Reads one cell as a "date" extra field. Cells are almost always a
 *  raw Excel serial number (or blank); the defensive fallback for an
 *  already-textual date (parsed with the platform's own Date parser)
 *  covers a hand-typed date string, which some sheets do have here and
 *  there, without throwing on it. Anything unparseable comes through
 *  blank rather than surfacing garbage. */
function extraFieldFromDateCell(cell: unknown): ExtraField {
  const raw = String(cell ?? "").trim();
  if (raw === "") {
    return { kind: "date", value: "" };
  }
  const serial = Number(raw);
  if (Number.isFinite(serial) && serial > 0) {
    return { kind: "date", value: excelSerialToIsoDate(serial) };
  }
  const parsed = new Date(raw);
  return {
    kind: "date",
    value: Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString().slice(0, 10),
  };
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
 *
 * `sourceId` scopes the synthesized id given to a source's *second and
 * later* occurrence of a repeated part_number (see partNumberOccurrences
 * below) — without it, two different sources that each happen to repeat
 * the same part_number the same number of times would synthesize the
 * exact same id (e.g. two sources both reusing "W11688994" a second
 * time would both get "sharepoint-w11688994-2"), so `parts.find(id)`
 * would silently resolve to whichever source's row happened to come
 * first in the merged list — editing "the 7300 Inventory row" would
 * silently open and overwrite Hadi Inventory's row instead. A first
 * occurrence doesn't need this: it's always deduped by part_number
 * across sources in mergeParts, so it never lands in the id-colliding
 * "standalone" list to begin with.
 */
export function mapTableRowsToRawParts(
  headers: string[],
  rows: unknown[][],
  columnMap: SharePointColumnMap = COLUMN_MAP,
  sourceId = "sharepoint",
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
  const extraKinds = new Map(
    extraIndices.map(({ header, idx }) => [idx, classifyExtraColumn(header, idx, rows)]),
  );

  // How many rows have already been seen for a given part_number, in
  // sheet order — a real sheet does sometimes list the same part_number
  // twice (see mergeParts in lib/sources/merge.ts, which relies on
  // RawPart.partNumberOccurrence to keep such rows as separate Parts
  // instead of colliding). Tracked regardless of whether this source has
  // an explicit id column, since it's also what lets updatePartInWorkbook
  // find the exact physical row an edit came from rather than always the
  // first row with that part_number. When there's no id column, the first
  // occurrence also keeps the plain part_number slug as its id (unchanged
  // from before this counter existed, so ordinary parts' ids and URLs
  // stay stable) — only the second and later occurrences get a numeric
  // suffix.
  const partNumberOccurrences = new Map<string, number>();

  return rows
    .filter((row) => String(row[partNumberIdx] ?? "").trim().length > 0)
    .map((row) => {
      const partNumber = String(row[partNumberIdx]).trim();
      const quantity = Number(row[quantityIdx]);

      const extraFields: ExtraFields = {};
      for (const { header, idx } of extraIndices) {
        const cell = row[idx];
        const classification = extraKinds.get(idx)!;
        if (classification.kind === "boolean") {
          const normalized = String(cell ?? "").trim().toUpperCase();
          extraFields[header] = { kind: "boolean", value: TRUE_CELL_VALUES.has(normalized) };
        } else if (classification.kind === "date") {
          extraFields[header] = extraFieldFromDateCell(cell);
        } else if (classification.kind === "select") {
          extraFields[header] = {
            kind: "select",
            value: String(cell ?? "").trim(),
            options: classification.options,
          };
        } else {
          extraFields[header] = { kind: "text", value: String(cell ?? "") };
        }
      }

      const occurrence = (partNumberOccurrences.get(partNumber) ?? 0) + 1;
      partNumberOccurrences.set(partNumber, occurrence);

      let id: string;
      if (idIdx !== -1) {
        id = String(row[idIdx]);
      } else {
        const slug = slugify(partNumber);
        // Only the 2nd-and-later-occurrence id needs the source baked
        // in (see this function's doc comment) — the 1st occurrence's
        // id is left exactly as it's always been, since changing it
        // would churn every ordinary (non-duplicated) part's id/URL for
        // no benefit.
        id = occurrence === 1 ? `sharepoint-${slug}` : `${sourceId}-${slug}-${occurrence}`;
      }

      return {
        id,
        part_number: partNumber,
        partNumberOccurrence: occurrence,
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
    let value: unknown = field.value;
    if (field.kind === "boolean") {
      value = field.value ? "YES" : "NO";
    } else if (field.kind === "date") {
      const iso = typeof field.value === "string" ? field.value.trim() : "";
      value = iso ? isoDateToExcelSerial(iso) : "";
    }
    assignments.push([header, value]);
  }
  return assignments;
}

/** Fills today's date into an "Entry Date"-shaped extra field when a
 *  brand-new part is added and nothing already supplied a value for
 *  it — see todayIsoDate. Deliberately only called from
 *  addPartToWorkbook, never updatePartInWorkbook: an existing part's
 *  entry date is the day it was first entered, not the day it was last
 *  edited, so a later edit must never stomp it back to "today". */
function withAutoEntryDate(headers: string[], fields: PartFields): PartFields {
  const header = headers.find((h) => /entry\s*date/i.test(String(h ?? "").trim()));
  if (!header) {
    return fields;
  }
  const existing = fields.extraFields?.[header];
  const hasValue = existing
    ? typeof existing.value === "string"
      ? existing.value.trim() !== ""
      : Boolean(existing.value)
    : false;
  if (hasValue) {
    return fields;
  }
  return {
    ...fields,
    extraFields: {
      ...fields.extraFields,
      [header]: { kind: "date", value: todayIsoDate() },
    },
  };
}

/** Finds this sheet's "UPN#" column, if it has one — matches "UPN",
 *  "UPN#", "upn #", etc. (the "sequential internal number" the user
 *  described — see withAutoNextUpn). */
function findUpnHeader(headers: string[]): string | undefined {
  return headers.find((h) => /^upn\s*#?$/i.test(String(h ?? "").trim()));
}

/** Walks a column from the bottom row up, looking for the last
 *  non-blank value, and returns one more than it (as a string, since
 *  extra-field values are always string | boolean — see ExtraField).
 *  Skips trailing blank cells — a still-blank template row, or a part
 *  like a hand-added test row that never got a UPN — rather than giving
 *  up the moment the very last row happens to be blank. Stops (returns
 *  undefined) the moment it hits a non-numeric value: "the next number
 *  after some free text" isn't something this can compute, and it's
 *  safer to leave the field for a human than to guess. */
function nextSequentialValue(rows: unknown[][], columnIndex: number): string | undefined {
  for (let i = rows.length - 1; i >= 0; i--) {
    const raw = String(rows[i]?.[columnIndex] ?? "").trim();
    if (raw === "") continue;
    const n = Number(raw);
    return Number.isFinite(n) ? String(Math.trunc(n) + 1) : undefined;
  }
  return undefined;
}

/** Fills a brand-new part's "UPN#" with one more than the last row's
 *  own UPN# — the shop's own "sequential internal number", picked up
 *  automatically so nobody has to look up the last one by hand. Like
 *  withAutoEntryDate, this only ever runs from addPartToWorkbook, never
 *  updatePartInWorkbook: an existing part keeps whatever UPN# it was
 *  given when it was first added, never renumbered on a later edit. A
 *  value the caller already supplied (there isn't one today — nothing
 *  in the "Add a part" form sets this — but the check costs nothing and
 *  means a future caller could) is left alone; a sheet with no UPN-like
 *  column, or whose last entry isn't a plain number, is left untouched
 *  too rather than guessing at a first value. */
function withAutoNextUpn(headers: string[], rows: unknown[][], fields: PartFields): PartFields {
  const header = findUpnHeader(headers);
  if (!header) {
    return fields;
  }
  const existing = fields.extraFields?.[header];
  const hasValue = existing
    ? typeof existing.value === "string"
      ? existing.value.trim() !== ""
      : Boolean(existing.value)
    : false;
  if (hasValue) {
    return fields;
  }
  const columnIndex = headerIndex(headers, header);
  const next = columnIndex === -1 ? undefined : nextSequentialValue(rows, columnIndex);
  if (next === undefined) {
    return fields;
  }
  return {
    ...fields,
    extraFields: {
      ...fields.extraFields,
      [header]: { kind: "text", value: next },
    },
  };
}

/**
 * Read-only preview of what addPartToWorkbook's withAutoNextUpn would
 * auto-fill this source's UPN#-shaped column with right now — for the
 * "Add a part" form to show the technician the next number up front,
 * before they've filled in anything else, rather than only learning it
 * after the part is already saved. Mirrors the Table/worksheet header
 * + row reading addPartToWorkbook does for the same purpose, but never
 * writes anything. Returns undefined for the same "leave it for a
 * human" cases withAutoNextUpn itself defers on: no UPN-shaped column
 * on this sheet, or its last entry isn't a plain number to build on.
 *
 * This is a genuine preview, not a reservation — nothing stops two
 * people opening the form at once from both seeing the same next
 * number, and whichever one saves second gets whatever
 * addPartToWorkbook computes fresh at that moment (skipping the
 * now-taken number only if the form still sent this preview value
 * along, which is unusual for this single-shop tool but not
 * impossible to hit).
 */
export async function peekNextUpn(
  client: Client,
  fileBase: string,
  tableOrWorksheetName: string,
): Promise<{ header: string; next: string } | undefined> {
  const tableBase = `${fileBase}/tables/${encodeURIComponent(tableOrWorksheetName)}`;
  const tableExists = await probeTableExists(client, tableBase, tableOrWorksheetName);

  if (tableExists) {
    const headerRange = await runStep(
      `Reading Table "${tableOrWorksheetName}"'s header row failed`,
      () => client.api(`${tableBase}/headerRowRange`).get(),
    );
    const headers: string[] = (headerRange.values?.[0] ?? []).map((v: unknown) => String(v));
    const header = findUpnHeader(headers);
    const columnIndex = header ? headerIndex(headers, header) : -1;
    if (!header || columnIndex === -1) {
      return undefined;
    }
    const rowsResponse = await runStep(
      `Reading Table "${tableOrWorksheetName}"'s rows failed`,
      () => client.api(`${tableBase}/rows`).get(),
    );
    const tableRows: unknown[][] = (rowsResponse.value ?? []).map(
      (row: { values: unknown[][] }) => row.values[0],
    );
    const next = nextSequentialValue(tableRows, columnIndex);
    return next === undefined ? undefined : { header, next };
  }

  const worksheetBase = `${fileBase}/worksheets/${encodeURIComponent(tableOrWorksheetName)}`;
  const usedRange = await runStep(
    `Reading worksheet "${tableOrWorksheetName}" failed`,
    () => client.api(`${worksheetBase}/usedRange`).get(),
  );
  const values: unknown[][] = usedRange.values ?? [];
  const headers = (values[0] ?? []).map((v: unknown) => String(v));
  const header = findUpnHeader(headers);
  const columnIndex = header ? headerIndex(headers, header) : -1;
  if (!header || columnIndex === -1) {
    return undefined;
  }
  const next = nextSequentialValue(values.slice(1), columnIndex);
  return next === undefined ? undefined : { header, next };
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
 *
 * `existingPartOccurrence` (1-based, defaults to 1) picks which row when
 * more than one shares `existingPartNumber` — see RawPart.partNumberOccurrence.
 * Without it, an edit to the *second* WP2163777 row, say, would silently
 * land on the *first* one instead, since matching by part_number alone
 * can't tell them apart.
 */
export async function updatePartInWorkbook(
  client: Client,
  fileBase: string,
  tableOrWorksheetName: string,
  columnMap: SharePointColumnMap,
  existingPartNumber: string,
  fields: PartFields,
  existingPartOccurrence: number = 1,
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
    const matchIdx = nthMatchIndex(
      tableRows,
      (row) => String(row.values[0]?.[partNumberIdx] ?? "").trim() === existingPartNumber,
      existingPartOccurrence,
    );
    const match = matchIdx === -1 ? undefined : tableRows[matchIdx];
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
  const matchOffset = nthMatchIndex(
    dataRows,
    (row) => String(row[partNumberIdx] ?? "").trim() === existingPartNumber,
    existingPartOccurrence,
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

    let effectiveFields = withAutoEntryDate(headers, fields);
    if (findUpnHeader(headers)) {
      // Only fetch the existing rows when this sheet actually has a
      // UPN-shaped column — most sources don't, and there's no reason
      // to pay for an extra Graph call on every part added to them.
      const rowsResponse = await runStep(
        `Reading Table "${tableOrWorksheetName}"'s rows failed`,
        () => client.api(`${tableBase}/rows`).get(),
      );
      const tableRows: unknown[][] = (rowsResponse.value ?? []).map(
        (row: { values: unknown[][] }) => row.values[0],
      );
      effectiveFields = withAutoNextUpn(headers, tableRows, effectiveFields);
    }
    applyFieldsToRow(newRow, headers, columnMap, effectiveFields);

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

  let effectiveFields = withAutoEntryDate(headers, fields);
  effectiveFields = withAutoNextUpn(headers, values.slice(1), effectiveFields);

  await writeFieldsToWorksheetRow(
    client,
    worksheetBase,
    tableOrWorksheetName,
    headers,
    startColumn0,
    newRowIndex0,
    columnMap,
    effectiveFields,
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
  const sourceId = identity?.id ?? "sharepoint";

  return {
    id: sourceId,
    label: identity?.label ?? `SharePoint workbook (${config.filePath})`,
    async fetchParts(): Promise<RawPart[]> {
      const { client, fileBase, filePath, fileUrl } = await connectToWorkbook(
        accessToken,
        config,
      );
      resolvedFileUrl = fileUrl;
      const { headers, rows } = await readWorkbookData(client, fileBase, filePath, config.tableName);

      return mapTableRowsToRawParts(headers, rows, config.columnMap ?? COLUMN_MAP, sourceId);
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
