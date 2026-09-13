import assert from "node:assert/strict";
import { test } from "node:test";
import type { Client } from "@microsoft/microsoft-graph-client";
import { mapTableRowsToRawParts, readWorkbookData } from "./sharepoint-excel-source";

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

function itemNotFoundError(): Error {
  return Object.assign(new Error("The requested resource doesn't exist."), {
    code: "itemNotFound",
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

test("mapTableRowsToRawParts: uses the configured id column when present", () => {
  const headers = ["RowId", "PartNumber", "QtyOnHand"];
  const rows = [["row-9", "WR17X11705", 14]];

  const parts = mapTableRowsToRawParts(headers, rows, {
    ...columnMap,
    id: "RowId",
  });

  assert.equal(parts[0].id, "row-9");
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
