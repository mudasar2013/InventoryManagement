import assert from "node:assert/strict";
import { test } from "node:test";
import type { Client } from "@microsoft/microsoft-graph-client";
import {
  isItemNotFoundError,
  mapTableRowsToRawParts,
  readWorkbookData,
  resolveFilePath,
} from "./sharepoint-excel-source";

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

test("resolveFilePath: resolves as-is when the path already works", async () => {
  const client = fakeClient({
    "/sites/site-id/drive/root:/Inventory.xlsx": async () => ({ id: "item-1" }),
  });

  const resolved = await resolveFilePath(client, "site-id", "Inventory.xlsx");

  assert.equal(resolved, "Inventory.xlsx");
});

test("resolveFilePath: falls back to stripping a leading \"Shared Documents/\" when the path as given 404s", async () => {
  const client = fakeClient({
    [`/sites/site-id/drive/root:/${encodeURI("Shared Documents/Inventory.xlsx")}`]: async () => {
      throw itemNotFoundError();
    },
    "/sites/site-id/drive/root:/Inventory.xlsx": async () => ({ id: "item-1" }),
  });

  const resolved = await resolveFilePath(client, "site-id", "Shared Documents/Inventory.xlsx");

  assert.equal(resolved, "Inventory.xlsx");
});

test("resolveFilePath: falls back to stripping a leading \"Documents/\" when the path as given 404s", async () => {
  const client = fakeClient({
    "/sites/site-id/drive/root:/Documents/Inventory.xlsx": async () => {
      throw itemNotFoundError();
    },
    "/sites/site-id/drive/root:/Inventory.xlsx": async () => ({ id: "item-1" }),
  });

  const resolved = await resolveFilePath(client, "site-id", "Documents/Inventory.xlsx");

  assert.equal(resolved, "Inventory.xlsx");
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
