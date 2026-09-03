import assert from "node:assert/strict";
import { test } from "node:test";
import { mapTableRowsToRawParts } from "./sharepoint-excel-source";

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
