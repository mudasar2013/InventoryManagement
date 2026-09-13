import assert from "node:assert/strict";
import { test } from "node:test";
import type { Client } from "@microsoft/microsoft-graph-client";
import {
  addPartToWorkbook,
  isItemNotFoundError,
  mapTableRowsToRawParts,
  readWorkbookData,
  resolveFilePath,
  updatePartInWorkbook,
} from "./sharepoint-excel-source";
import type { PartFields, SharePointColumnMap } from "./sharepoint-excel-source";

/**
 * A minimal stand-in for the Graph SDK's Client — readWorkbookData only
 * ever calls `client.api(path).get()`, so a map of path -> handler is
 * enough to drive it through both the Table path and the worksheet
 * fallback without a live SharePoint connection.
 */
function fakeClient(handlers: Record<string, () => Promise<unknown>>): Client {
  return {
    api: (path: string) => ({
      get: () => {
        const handler = handlers[path];
        if (!handler) {
          throw new Error(`Test fake received an unexpected Graph path: ${path}`);
        }
        return handler();
      },
    }),
  } as unknown as Client;
}

/**
 * Like fakeClient, but also records PATCH/POST calls so a write-path
 * test can assert exactly which cell/row was written and with what
 * body, without a live SharePoint connection. GET handlers are plain
 * `() => Promise<unknown>` (same as fakeClient); PATCH/POST handlers
 * receive the request body and default to `async () => ({})` when not
 * given, since most write tests only care that the right call was
 * made, not what it returns.
 */
function fakeWriteClient(
  getHandlers: Record<string, () => Promise<unknown>>,
): { client: Client; patches: { path: string; body: unknown }[]; posts: { path: string; body: unknown }[] } {
  const patches: { path: string; body: unknown }[] = [];
  const posts: { path: string; body: unknown }[] = [];
  const client = {
    api: (path: string) => ({
      get: () => {
        const handler = getHandlers[path];
        if (!handler) {
          throw new Error(`Test fake received an unexpected GET: ${path}`);
        }
        return handler();
      },
      patch: (body: unknown) => {
        patches.push({ path, body });
        return Promise.resolve({});
      },
      post: (body: unknown) => {
        posts.push({ path, body });
        return Promise.resolve({});
      },
    }),
  } as unknown as Client;
  return { client, patches, posts };
}

/** Independent re-implementation of the standard Excel-serial-date
 *  formula (day 0 = 1899-12-30), used only to compute expected values
 *  for the date-classification/round-trip tests below — deliberately
 *  not imported from sharepoint-excel-source.ts, since these are
 *  internal helpers and the point is to check the module's behavior
 *  against the well-known algorithm, not against itself. */
const EXCEL_EPOCH_UTC_MS_FOR_TEST = Date.UTC(1899, 11, 30);
function isoFromExcelSerialForTest(serial: number): string {
  return new Date(EXCEL_EPOCH_UTC_MS_FOR_TEST + serial * 86400000).toISOString().slice(0, 10);
}
function excelSerialFromIsoForTest(iso: string): number {
  const [year, month, day] = iso.split("-").map(Number);
  return Math.round((Date.UTC(year, month - 1, day) - EXCEL_EPOCH_UTC_MS_FOR_TEST) / 86400000);
}

function itemNotFoundError(): Error {
  return Object.assign(new Error("The requested resource doesn't exist."), {
    code: "itemNotFound",
  });
}

/** The workbook Tables endpoint has been observed returning this exact
 *  casing ("ItemNotFound", capital I) for a missing Table, unlike the
 *  drive/file endpoints which use "itemNotFound" — see the doc comment
 *  on isItemNotFoundError. */
function capitalizedItemNotFoundError(): Error {
  return Object.assign(new Error("The requested resource doesn't exist."), {
    code: "ItemNotFound",
  });
}

const columnMap = {
  part_number: "PartNumber",
  description: "Description",
  bin_location: "BinLocation",
  quantity_on_hand: "QtyOnHand",
  id: undefined as string | undefined,
};

test("mapTableRowsToRawParts: maps rows by header name, not column position", () => {
  // Columns deliberately out of the "expected" order — this is the whole
  // point of mapping by header text instead of a fixed index.
  const headers = ["QtyOnHand", "PartNumber", "BinLocation", "Description"];
  const rows = [[14, "WR17X11705", "A-12-04", "Water filter cartridge"]];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.equal(parts.length, 1);
  assert.equal(parts[0].part_number, "WR17X11705");
  assert.equal(parts[0].quantity_on_hand, 14);
  assert.equal(parts[0].bin_location, "A-12-04");
  assert.equal(parts[0].description, "Water filter cartridge");
});

test("mapTableRowsToRawParts: synthesizes a stable id from part_number when no id column is configured", () => {
  const headers = ["PartNumber", "QtyOnHand"];
  const rows = [["WR17X 11705!", 14]];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.equal(parts[0].id, "sharepoint-wr17x-11705");
});

test("mapTableRowsToRawParts: a duplicate part_number gets a distinct, suffixed id and an incrementing occurrence", () => {
  const headers = ["PartNumber", "QtyOnHand"];
  const rows = [
    ["WP2163777", 1],
    ["OTHER123", 5],
    ["WP2163777", 1],
    ["WP2163777", 0],
  ];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.equal(parts.length, 4);
  assert.equal(parts[0].id, "sharepoint-wp2163777");
  assert.equal(parts[0].partNumberOccurrence, 1);
  assert.equal(parts[1].id, "sharepoint-other123");
  assert.equal(parts[1].partNumberOccurrence, 1);
  assert.equal(parts[2].id, "sharepoint-wp2163777-2");
  assert.equal(parts[2].partNumberOccurrence, 2);
  assert.equal(parts[3].id, "sharepoint-wp2163777-3");
  assert.equal(parts[3].partNumberOccurrence, 3);
});

test("mapTableRowsToRawParts: uses the configured id column when present", () => {
  const headers = ["RowId", "PartNumber", "QtyOnHand"];
  const rows = [["row-9", "WR17X11705", 14]];

  const parts = mapTableRowsToRawParts(headers, rows, {
    ...columnMap,
    id: "RowId",
  });

  assert.equal(parts[0].id, "row-9");
});

test("mapTableRowsToRawParts: still tracks partNumberOccurrence even when an explicit id column is configured", () => {
  const headers = ["RowId", "PartNumber", "QtyOnHand"];
  const rows = [
    ["row-1", "WP2163777", 1],
    ["row-2", "WP2163777", 0],
  ];

  const parts = mapTableRowsToRawParts(headers, rows, { ...columnMap, id: "RowId" });

  assert.equal(parts[0].id, "row-1");
  assert.equal(parts[0].partNumberOccurrence, 1);
  assert.equal(parts[1].id, "row-2");
  assert.equal(parts[1].partNumberOccurrence, 2);
});

test("mapTableRowsToRawParts: skips rows with a blank part number", () => {
  const headers = ["PartNumber", "QtyOnHand"];
  const rows = [
    ["WR17X11705", 14],
    ["", 3],
    ["   ", 3],
  ];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.equal(parts.length, 1);
});

test("mapTableRowsToRawParts: a non-numeric quantity cell falls back to 0 instead of NaN", () => {
  const headers = ["PartNumber", "QtyOnHand"];
  const rows = [["WR17X11705", "N/A"]];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.equal(parts[0].quantity_on_hand, 0);
});

test("mapTableRowsToRawParts: a header containing \"date\" is read as kind date, converting the Excel serial to an ISO string", () => {
  const headers = ["PartNumber", "QtyOnHand", "Entry Date"];
  const rows = [["WR17X11705", 5, 45866]];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.deepEqual(parts[0].extraFields?.["Entry Date"], {
    kind: "date",
    value: isoFromExcelSerialForTest(45866),
  });
});

test("mapTableRowsToRawParts: a blank date cell reads as an empty date value, not NaN or the raw blank", () => {
  const headers = ["PartNumber", "QtyOnHand", "Date Rcvd"];
  const rows = [["WR17X11705", 5, ""]];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.deepEqual(parts[0].extraFields?.["Date Rcvd"], { kind: "date", value: "" });
});

test("mapTableRowsToRawParts: a header naming itself \"(Yes/No)\" is boolean even when every cell seen so far is blank", () => {
  const headers = ["PartNumber", "QtyOnHand", "Ebay Ready (Yes/No)"];
  const rows = [["WR17X11705", 5, ""]];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.deepEqual(parts[0].extraFields?.["Ebay Ready (Yes/No)"], {
    kind: "boolean",
    value: false,
  });
});

test("mapTableRowsToRawParts: \"Condition\" is a select field with fixed options, and an out-of-list value is preserved as-is", () => {
  const headers = ["PartNumber", "QtyOnHand", "Condition"];
  const rows = [["WR17X11705", 5, "Used But Working"]];

  const parts = mapTableRowsToRawParts(headers, rows, columnMap);

  assert.deepEqual(parts[0].extraFields?.["Condition"], {
    kind: "select",
    value: "Used But Working",
    options: ["New", "Used", "OpenBox", "Used/Working", "Other"],
  });
});

test("mapTableRowsToRawParts: throws a clear error when a required column is missing", () => {
  const headers = ["SKU", "Count"]; // doesn't match columnMap at all
  const rows: unknown[][] = [];

  assert.throws(
    () => mapTableRowsToRawParts(headers, rows, columnMap),
    /missing an expected column/,
  );
});

test("readWorkbookData: reads from an Excel Table when one by that name exists", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const client = fakeClient({
    [`${fileBase}/tables/Parts`]: async () => ({ id: "table-1" }),
    [`${fileBase}/tables/Parts/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand"]],
    }),
    [`${fileBase}/tables/Parts/rows`]: async () => ({
      value: [{ values: [["WR17X11705", 5]] }],
    }),
  });

  const result = await readWorkbookData(client, fileBase, "Inventory.xlsx", "Parts");

  assert.deepEqual(result.headers, ["PartNumber", "QtyOnHand"]);
  assert.deepEqual(result.rows, [["WR17X11705", 5]]);
});

test("readWorkbookData: falls back to a worksheet's used range when no Table by that name exists", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const client = fakeClient({
    [`${fileBase}/tables/Sheet1`]: async () => {
      throw itemNotFoundError();
    },
    [`${fileBase}/worksheets/Sheet1/usedRange`]: async () => ({
      values: [
        ["PartNumber", "QtyOnHand"],
        ["WR17X11705", 5],
      ],
    }),
  });

  const result = await readWorkbookData(client, fileBase, "Inventory.xlsx", "Sheet1");

  assert.deepEqual(result.headers, ["PartNumber", "QtyOnHand"]);
  assert.deepEqual(result.rows, [["WR17X11705", 5]]);
});

test("readWorkbookData: still falls back to the worksheet when the Table probe 404s with capitalized \"ItemNotFound\"", async () => {
  // The workbook Tables endpoint has been observed returning this
  // casing for a missing Table, unlike the "itemNotFound" the drive/file
  // endpoints use for the same situation — this must not be treated as
  // a hard failure that skips the worksheet fallback.
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const client = fakeClient({
    [`${fileBase}/tables/OfficeInventory`]: async () => {
      throw capitalizedItemNotFoundError();
    },
    [`${fileBase}/worksheets/OfficeInventory/usedRange`]: async () => ({
      values: [
        ["PartNumber", "QtyOnHand"],
        ["WR17X11705", 5],
      ],
    }),
  });

  const result = await readWorkbookData(client, fileBase, "Inventory.xlsx", "OfficeInventory");

  assert.deepEqual(result.headers, ["PartNumber", "QtyOnHand"]);
  assert.deepEqual(result.rows, [["WR17X11705", 5]]);
});

test("isItemNotFoundError: matches regardless of case", () => {
  assert.equal(isItemNotFoundError(itemNotFoundError()), true);
  assert.equal(isItemNotFoundError(capitalizedItemNotFoundError()), true);
  assert.equal(
    isItemNotFoundError(Object.assign(new Error("nope"), { code: "Forbidden" })),
    false,
  );
  assert.equal(isItemNotFoundError(new Error("no code at all")), false);
});

test("readWorkbookData: a non-itemNotFound error on the Table probe propagates instead of falling back", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  let worksheetWasCalled = false;
  const client = fakeClient({
    [`${fileBase}/tables/Parts`]: async () => {
      throw Object.assign(new Error("Forbidden"), { code: "Forbidden", statusCode: 403 });
    },
    [`${fileBase}/worksheets/Parts/usedRange`]: async () => {
      worksheetWasCalled = true;
      return { values: [] };
    },
  });

  await assert.rejects(() => readWorkbookData(client, fileBase, "Inventory.xlsx", "Parts"));
  assert.equal(worksheetWasCalled, false);
});

test("readWorkbookData: names both the file path and the table/worksheet name when neither resolves", async () => {
  const fileBase = "/sites/site-id/drive/root:/Wrong.xlsx:/workbook";
  const client = fakeClient({
    [`${fileBase}/tables/Parts`]: async () => {
      throw itemNotFoundError();
    },
    [`${fileBase}/worksheets/Parts/usedRange`]: async () => {
      throw itemNotFoundError();
    },
  });

  await assert.rejects(
    () => readWorkbookData(client, fileBase, "Wrong.xlsx", "Parts"),
    (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /Wrong\.xlsx/);
      assert.match(error.message, /Parts/);
      return true;
    },
  );
});

test("resolveFilePath: resolves as-is when the path already works, and returns the item's webUrl", async () => {
  const client = fakeClient({
    "/sites/site-id/drive/root:/Inventory.xlsx": async () => ({
      id: "item-1",
      webUrl: "https://contoso.sharepoint.com/sites/Foo/Inventory.xlsx",
    }),
  });

  const resolved = await resolveFilePath(client, "site-id", "Inventory.xlsx");

  assert.equal(resolved.path, "Inventory.xlsx");
  assert.equal(resolved.webUrl, "https://contoso.sharepoint.com/sites/Foo/Inventory.xlsx");
});

test("resolveFilePath: falls back to stripping a leading \"Shared Documents/\" when the path as given 404s", async () => {
  const client = fakeClient({
    [`/sites/site-id/drive/root:/${encodeURI("Shared Documents/Inventory.xlsx")}`]: async () => {
      throw itemNotFoundError();
    },
    "/sites/site-id/drive/root:/Inventory.xlsx": async () => ({ id: "item-1" }),
  });

  const resolved = await resolveFilePath(client, "site-id", "Shared Documents/Inventory.xlsx");

  assert.equal(resolved.path, "Inventory.xlsx");
});

test("resolveFilePath: falls back to stripping a leading \"Documents/\" when the path as given 404s", async () => {
  const client = fakeClient({
    "/sites/site-id/drive/root:/Documents/Inventory.xlsx": async () => {
      throw itemNotFoundError();
    },
    "/sites/site-id/drive/root:/Inventory.xlsx": async () => ({ id: "item-1" }),
  });

  const resolved = await resolveFilePath(client, "site-id", "Documents/Inventory.xlsx");

  assert.equal(resolved.path, "Inventory.xlsx");
});

test("resolveFilePath: leaves webUrl undefined when the resolved item's response doesn't include one", async () => {
  const client = fakeClient({
    "/sites/site-id/drive/root:/Inventory.xlsx": async () => ({ id: "item-1" }),
  });

  const resolved = await resolveFilePath(client, "site-id", "Inventory.xlsx");

  assert.equal(resolved.webUrl, undefined);
});

test("resolveFilePath: names both candidates it tried when neither resolves", async () => {
  const client = fakeClient({
    [`/sites/site-id/drive/root:/${encodeURI("Shared Documents/Missing.xlsx")}`]: async () => {
      throw itemNotFoundError();
    },
    "/sites/site-id/drive/root:/Missing.xlsx": async () => {
      throw itemNotFoundError();
    },
  });

  await assert.rejects(
    () => resolveFilePath(client, "site-id", "Shared Documents/Missing.xlsx"),
    (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /Shared Documents\/Missing\.xlsx/);
      assert.match(error.message, /Missing\.xlsx/);
      return true;
    },
  );
});

test("resolveFilePath: a non-itemNotFound error on the first candidate propagates without trying a second", async () => {
  let secondCandidateWasCalled = false;
  const client = fakeClient({
    [`/sites/site-id/drive/root:/${encodeURI("Shared Documents/Inventory.xlsx")}`]: async () => {
      throw Object.assign(new Error("Forbidden"), { code: "Forbidden", statusCode: 403 });
    },
    "/sites/site-id/drive/root:/Inventory.xlsx": async () => {
      secondCandidateWasCalled = true;
      return { id: "item-1" };
    },
  });

  await assert.rejects(() =>
    resolveFilePath(client, "site-id", "Shared Documents/Inventory.xlsx"),
  );
  assert.equal(secondCandidateWasCalled, false);
});

test("resolveFilePath: a path with no library-name prefix has only one candidate to try", async () => {
  const client = fakeClient({
    [`/sites/site-id/drive/root:/${encodeURI("Team Channel/Inventory.xlsx")}`]: async () => {
      throw itemNotFoundError();
    },
  });

  await assert.rejects(
    () => resolveFilePath(client, "site-id", "Team Channel/Inventory.xlsx"),
    (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /Team Channel\/Inventory\.xlsx/);
      assert.doesNotMatch(error.message, /also tried/);
      return true;
    },
  );
});

const tableColumnMap: SharePointColumnMap = {
  part_number: "PartNumber",
  quantity_on_hand: "QtyOnHand",
};

const scatteredColumnMap: SharePointColumnMap = {
  part_number: "PartNumber",
  quantity_on_hand: "Quantity",
  description: "PartName",
  bin_location: "Location",
};

test("updatePartInWorkbook: Table path PATCHes the matched row, updating only mapped columns", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, patches } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand", "Extra"]],
    }),
    [`${tableBase}/rows`]: async () => ({
      value: [
        { index: 0, values: [["OLD123", 3, "keep-me"]] },
        { index: 1, values: [["WR17X11705", 5, "other"]] },
      ],
    }),
  });

  const fields: PartFields = {
    part_number: "WR17X11705",
    description: "",
    bin_location: "",
    quantity_on_hand: 9,
  };
  await updatePartInWorkbook(client, fileBase, "Parts", tableColumnMap, "WR17X11705", fields);

  assert.equal(patches.length, 1);
  assert.equal(patches[0].path, `${tableBase}/rows/itemAt(index=1)`);
  assert.deepEqual(patches[0].body, { values: [["WR17X11705", 9, "other"]] });
});

test("updatePartInWorkbook: Table path picks the Nth row when existingPartOccurrence disambiguates a duplicate part number", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, patches } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand"]],
    }),
    [`${tableBase}/rows`]: async () => ({
      value: [
        { index: 0, values: [["WP2163777", 1]] },
        { index: 1, values: [["WP2163777", 5]] },
      ],
    }),
  });

  const fields: PartFields = {
    part_number: "WP2163777",
    description: "",
    bin_location: "",
    quantity_on_hand: 9,
  };
  // existingPartOccurrence 2 must land on the *second* WP2163777 row
  // (table index 1), not the first one a plain part_number match would
  // find — this is the whole point of the occurrence parameter.
  await updatePartInWorkbook(client, fileBase, "Parts", tableColumnMap, "WP2163777", fields, 2);

  assert.equal(patches.length, 1);
  assert.equal(patches[0].path, `${tableBase}/rows/itemAt(index=1)`);
  assert.deepEqual(patches[0].body, { values: [["WP2163777", 9]] });
});

test("updatePartInWorkbook: existingPartOccurrence defaults to 1, matching the first row", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, patches } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand"]],
    }),
    [`${tableBase}/rows`]: async () => ({
      value: [
        { index: 0, values: [["WP2163777", 1]] },
        { index: 1, values: [["WP2163777", 5]] },
      ],
    }),
  });

  await updatePartInWorkbook(client, fileBase, "Parts", tableColumnMap, "WP2163777", {
    part_number: "WP2163777",
    description: "",
    bin_location: "",
    quantity_on_hand: 9,
  });

  assert.equal(patches[0].path, `${tableBase}/rows/itemAt(index=0)`);
});

test("updatePartInWorkbook: Table path throws a clear error when existingPartOccurrence asks for a row that doesn't exist", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({ values: [["PartNumber", "QtyOnHand"]] }),
    [`${tableBase}/rows`]: async () => ({ value: [{ index: 0, values: [["WP2163777", 1]] }] }),
  });

  await assert.rejects(
    () =>
      updatePartInWorkbook(
        client,
        fileBase,
        "Parts",
        tableColumnMap,
        "WP2163777",
        { part_number: "WP2163777", description: "", bin_location: "", quantity_on_hand: 9 },
        2,
      ),
    /no row with part number "WP2163777"/,
  );
});

test("updatePartInWorkbook: worksheet path picks the Nth row when existingPartOccurrence disambiguates a duplicate part number", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Sheet1`;
  const worksheetBase = `${fileBase}/worksheets/Sheet1`;
  const { client, patches } = fakeWriteClient({
    [tableBase]: async () => {
      throw itemNotFoundError();
    },
    [`${worksheetBase}/usedRange`]: async () => ({
      rowIndex: 0,
      columnIndex: 0,
      values: [
        ["PartNumber", "QtyOnHand"],
        ["WP2163777", 1],
        ["WP2163777", 5],
      ],
    }),
  });

  await updatePartInWorkbook(
    client,
    fileBase,
    "Sheet1",
    tableColumnMap,
    "WP2163777",
    { part_number: "WP2163777", description: "", bin_location: "", quantity_on_hand: 9 },
    2,
  );

  // Second data row (0-based offset 1) -> absolute row 0 (usedRange) + 1
  // (header) + 1 (offset) = 2 (0-based) = row 3.
  const byAddress = new Map(
    patches.map((p) => [p.path, (p.body as { values: unknown[][] }).values[0][0]]),
  );
  assert.equal(byAddress.get(`${worksheetBase}/range(address='B3')`), 9);
});

test("updatePartInWorkbook: Table path throws a clear error when no row matches the part number", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({ values: [["PartNumber", "QtyOnHand"]] }),
    [`${tableBase}/rows`]: async () => ({ value: [{ index: 0, values: [["OLD123", 3]] }] }),
  });

  await assert.rejects(
    () =>
      updatePartInWorkbook(client, fileBase, "Parts", tableColumnMap, "WR17X11705", {
        part_number: "WR17X11705",
        description: "",
        bin_location: "",
        quantity_on_hand: 9,
      }),
    /no row with part number "WR17X11705"/,
  );
});

test("updatePartInWorkbook: worksheet path writes each mapped, scattered column as its own cell PATCH", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Sheet1`;
  const worksheetBase = `${fileBase}/worksheets/Sheet1`;
  const { client, patches } = fakeWriteClient({
    [tableBase]: async () => {
      throw itemNotFoundError();
    },
    [`${worksheetBase}/usedRange`]: async () => ({
      rowIndex: 2,
      columnIndex: 1,
      values: [
        ["Location", "PartName", "PartNumber", "Quantity"],
        ["A-1", "Water filter", "OLD999", 3],
        ["A-2", "Ice maker valve", "WR17X11705", 5],
      ],
    }),
  });

  const fields: PartFields = {
    part_number: "WR17X11705",
    description: "New desc",
    bin_location: "A-9",
    quantity_on_hand: 12,
  };
  await updatePartInWorkbook(
    client,
    fileBase,
    "Sheet1",
    scatteredColumnMap,
    "WR17X11705",
    fields,
  );

  // usedRange starts at 0-based row 2 (row 3), column 1 (column B); the
  // matched data row is the second one (0-based offset 1) — so absolute
  // row = 2 + 1 (header) + 1 (offset) = 4 (0-based) = row 5. Headers are
  // Location(B), PartName(C), PartNumber(D), Quantity(E) in that column
  // order, so each field lands on its own row-5 cell.
  const byAddress = new Map(
    patches.map((p) => [p.path, (p.body as { values: unknown[][] }).values[0][0]]),
  );
  assert.equal(byAddress.get(`${worksheetBase}/range(address='D5')`), "WR17X11705");
  assert.equal(byAddress.get(`${worksheetBase}/range(address='C5')`), "New desc");
  assert.equal(byAddress.get(`${worksheetBase}/range(address='B5')`), "A-9");
  assert.equal(byAddress.get(`${worksheetBase}/range(address='E5')`), 12);
  assert.equal(patches.length, 4);
});

test("updatePartInWorkbook: worksheet path throws a clear error when no row matches the part number", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Sheet1`;
  const worksheetBase = `${fileBase}/worksheets/Sheet1`;
  const { client } = fakeWriteClient({
    [tableBase]: async () => {
      throw itemNotFoundError();
    },
    [`${worksheetBase}/usedRange`]: async () => ({
      rowIndex: 0,
      columnIndex: 0,
      values: [
        ["PartNumber", "Quantity"],
        ["OLD999", 3],
      ],
    }),
  });

  await assert.rejects(
    () =>
      updatePartInWorkbook(client, fileBase, "Sheet1", scatteredColumnMap, "WR17X11705", {
        part_number: "WR17X11705",
        description: "",
        bin_location: "",
        quantity_on_hand: 12,
      }),
    /no row with part number "WR17X11705"/,
  );
});

test("updatePartInWorkbook: Table path writes a date extra field back as an Excel serial number, not the ISO string", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, patches } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand", "Entry Date"]],
    }),
    [`${tableBase}/rows`]: async () => ({
      value: [{ index: 0, values: [["WR17X11705", 5, 45866]] }],
    }),
  });

  await updatePartInWorkbook(client, fileBase, "Parts", tableColumnMap, "WR17X11705", {
    part_number: "WR17X11705",
    description: "",
    bin_location: "",
    quantity_on_hand: 5,
    extraFields: { "Entry Date": { kind: "date", value: "2026-08-13" } },
  });

  assert.deepEqual(patches[0].body, {
    values: [["WR17X11705", 5, excelSerialFromIsoForTest("2026-08-13")]],
  });
});

test("addPartToWorkbook: Table path POSTs a full-width row with mapped fields placed and everything else blank", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, posts } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand", "Extra"]],
    }),
  });

  await addPartToWorkbook(client, fileBase, "Parts", tableColumnMap, {
    part_number: "NEW123",
    description: "",
    bin_location: "",
    quantity_on_hand: 4,
  });

  assert.equal(posts.length, 1);
  assert.equal(posts[0].path, `${tableBase}/rows/add`);
  assert.deepEqual(posts[0].body, { values: [["NEW123", 4, ""]] });
});

test("addPartToWorkbook: Table path auto-fills an untouched \"Entry Date\" column with today's date", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, posts } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand", "Entry Date"]],
    }),
  });

  await addPartToWorkbook(client, fileBase, "Parts", tableColumnMap, {
    part_number: "NEW123",
    description: "",
    bin_location: "",
    quantity_on_hand: 4,
  });

  const now = new Date();
  const todayIso =
    `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-` +
    `${String(now.getDate()).padStart(2, "0")}`;

  assert.equal(posts.length, 1);
  assert.deepEqual(posts[0].body, {
    values: [["NEW123", 4, excelSerialFromIsoForTest(todayIso)]],
  });
});

test("addPartToWorkbook: does not override an \"Entry Date\" the caller already supplied", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, posts } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand", "Entry Date"]],
    }),
  });

  await addPartToWorkbook(client, fileBase, "Parts", tableColumnMap, {
    part_number: "NEW123",
    description: "",
    bin_location: "",
    quantity_on_hand: 4,
    extraFields: { "Entry Date": { kind: "date", value: "2026-08-13" } },
  });

  assert.deepEqual(posts[0].body, {
    values: [["NEW123", 4, excelSerialFromIsoForTest("2026-08-13")]],
  });
});

test("addPartToWorkbook: worksheet path writes a new row just past the current used range", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Sheet1`;
  const worksheetBase = `${fileBase}/worksheets/Sheet1`;
  const { client, patches } = fakeWriteClient({
    [tableBase]: async () => {
      throw itemNotFoundError();
    },
    [`${worksheetBase}/usedRange`]: async () => ({
      rowIndex: 0,
      columnIndex: 0,
      values: [
        ["PartNumber", "QtyOnHand"],
        ["OLD1", 3],
      ],
    }),
  });

  await addPartToWorkbook(client, fileBase, "Sheet1", tableColumnMap, {
    part_number: "NEW123",
    description: "",
    bin_location: "",
    quantity_on_hand: 4,
  });

  const byAddress = new Map(
    patches.map((p) => [p.path, (p.body as { values: unknown[][] }).values[0][0]]),
  );
  assert.equal(byAddress.get(`${worksheetBase}/range(address='A3')`), "NEW123");
  assert.equal(byAddress.get(`${worksheetBase}/range(address='B3')`), 4);
  assert.equal(patches.length, 2);
});

test("addPartToWorkbook: Table path auto-fills UPN# with one more than the last row's value", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, posts } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand", "UPN#"]],
    }),
    [`${tableBase}/rows`]: async () => ({
      value: [
        { index: 0, values: [["OLD1", 3, "594"]] },
        { index: 1, values: [["OLD2", 1, "595"]] },
      ],
    }),
  });

  await addPartToWorkbook(client, fileBase, "Parts", tableColumnMap, {
    part_number: "NEW123",
    description: "",
    bin_location: "",
    quantity_on_hand: 4,
  });

  assert.equal(posts.length, 1);
  assert.deepEqual(posts[0].body, { values: [["NEW123", 4, "596"]] });
});

test("addPartToWorkbook: does not override a UPN# the caller already supplied", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, posts } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand", "UPN#"]],
    }),
    // addPartToWorkbook still fetches the existing rows whenever the
    // sheet has a UPN-shaped column (it doesn't know yet whether the
    // caller already supplied a value), but withAutoNextUpn must leave
    // an explicit value alone rather than overwriting it.
    [`${tableBase}/rows`]: async () => ({
      value: [{ index: 0, values: [["OLD1", 3, "1"]] }],
    }),
  });

  await addPartToWorkbook(client, fileBase, "Parts", tableColumnMap, {
    part_number: "NEW123",
    description: "",
    bin_location: "",
    quantity_on_hand: 4,
    extraFields: { "UPN#": { kind: "text", value: "999" } },
  });

  assert.deepEqual(posts[0].body, { values: [["NEW123", 4, "999"]] });
});

test("addPartToWorkbook: Table path leaves UPN# blank when the column has no prior numeric value to build on", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Parts`;
  const { client, posts } = fakeWriteClient({
    [tableBase]: async () => ({ id: "table-1" }),
    [`${tableBase}/headerRowRange`]: async () => ({
      values: [["PartNumber", "QtyOnHand", "UPN#"]],
    }),
    [`${tableBase}/rows`]: async () => ({ value: [] }),
  });

  await addPartToWorkbook(client, fileBase, "Parts", tableColumnMap, {
    part_number: "NEW123",
    description: "",
    bin_location: "",
    quantity_on_hand: 4,
  });

  assert.deepEqual(posts[0].body, { values: [["NEW123", 4, ""]] });
});

test("addPartToWorkbook: worksheet path auto-fills UPN# with one more than the last row's value", async () => {
  const fileBase = "/sites/site-id/drive/root:/Inventory.xlsx:/workbook";
  const tableBase = `${fileBase}/tables/Sheet1`;
  const worksheetBase = `${fileBase}/worksheets/Sheet1`;
  const { client, patches } = fakeWriteClient({
    [tableBase]: async () => {
      throw itemNotFoundError();
    },
    [`${worksheetBase}/usedRange`]: async () => ({
      rowIndex: 0,
      columnIndex: 0,
      values: [
        ["PartNumber", "QtyOnHand", "UPN#"],
        ["OLD1", 3, "594"],
        ["OLD2", 1, "595"],
      ],
    }),
  });

  await addPartToWorkbook(client, fileBase, "Sheet1", tableColumnMap, {
    part_number: "NEW123",
    description: "",
    bin_location: "",
    quantity_on_hand: 4,
  });

  const byAddress = new Map(
    patches.map((p) => [p.path, (p.body as { values: unknown[][] }).values[0][0]]),
  );
  assert.equal(byAddress.get(`${worksheetBase}/range(address='C4')`), "596");
});
