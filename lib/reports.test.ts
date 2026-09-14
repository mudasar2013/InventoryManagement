import assert from "node:assert/strict";
import { test } from "node:test";
import {
  discoverExtraColumns,
  extraColumnValues,
  matchesNumericFilter,
  matchesOptionFilter,
  matchesTextFilter,
  toCsv,
} from "./reports";
import type { Part } from "./types";

function part(overrides: Partial<Part> = {}): Part {
  return {
    id: "prt-1",
    part_number: "WR17X11705",
    description: "Water filter cartridge",
    bin_location: "A-12-04",
    quantity_on_hand: 5,
    status: "In Stock",
    ...overrides,
  };
}

test("discoverExtraColumns: finds every distinct header across the catalog, sorted alphabetically", () => {
  const columns = discoverExtraColumns([
    part({ extraFields: { Condition: { kind: "select", value: "New" } } }),
    part({ extraFields: { "UPC#": { kind: "text", value: "123" } } }),
  ]);

  assert.deepEqual(
    columns.map((c) => c.header),
    ["Condition", "UPC#"],
  );
});

test("discoverExtraColumns: a header missing from some parts doesn't affect the others' classification", () => {
  const columns = discoverExtraColumns([
    part({ extraFields: { Condition: { kind: "select", value: "New" } } }),
    part({ extraFields: {} }),
  ]);

  assert.deepEqual(columns, [{ header: "Condition", kind: "select" }]);
});

test("discoverExtraColumns: picks whichever kind was seen most often for a header sources disagree on", () => {
  const columns = discoverExtraColumns([
    part({ extraFields: { Notes: { kind: "text", value: "a" } } }),
    part({ extraFields: { Notes: { kind: "text", value: "b" } } }),
    part({ extraFields: { Notes: { kind: "boolean", value: true } } }),
  ]);

  assert.deepEqual(columns, [{ header: "Notes", kind: "text" }]);
});

test("extraColumnValues: boolean cells become Yes/No, deduped and sorted", () => {
  const values = extraColumnValues(
    [
      part({ extraFields: { "Ebay Ready (Yes/No)": { kind: "boolean", value: true } } }),
      part({ extraFields: { "Ebay Ready (Yes/No)": { kind: "boolean", value: false } } }),
      part({ extraFields: { "Ebay Ready (Yes/No)": { kind: "boolean", value: true } } }),
    ],
    "Ebay Ready (Yes/No)",
  );

  assert.deepEqual(values, ["No", "Yes"]);
});

test("extraColumnValues: a part with no such column is skipped, not counted as a blank value", () => {
  const values = extraColumnValues(
    [
      part({ extraFields: { Condition: { kind: "select", value: "New" } } }),
      part({ extraFields: {} }),
    ],
    "Condition",
  );

  assert.deepEqual(values, ["New"]);
});

test("matchesTextFilter: empty query always matches; otherwise case-insensitive substring", () => {
  assert.equal(matchesTextFilter("Water filter cartridge", ""), true);
  assert.equal(matchesTextFilter("Water filter cartridge", "FILTER"), true);
  assert.equal(matchesTextFilter("Water filter cartridge", "motor"), false);
  assert.equal(matchesTextFilter(undefined, "motor"), false);
  assert.equal(matchesTextFilter(undefined, ""), true);
});

test("matchesOptionFilter: empty selection (Any) always matches; otherwise exact membership", () => {
  assert.equal(matchesOptionFilter(["Hadi Inventory", "7300 Inventory"], ""), true);
  assert.equal(matchesOptionFilter(["Hadi Inventory", "7300 Inventory"], "7300 Inventory"), true);
  assert.equal(matchesOptionFilter(["Hadi Inventory"], "7300 Inventory"), false);
});

test("matchesNumericFilter: unset bounds always match; each bound is inclusive", () => {
  assert.equal(matchesNumericFilter(5, {}), true);
  assert.equal(matchesNumericFilter(5, { min: 5 }), true);
  assert.equal(matchesNumericFilter(4, { min: 5 }), false);
  assert.equal(matchesNumericFilter(5, { max: 5 }), true);
  assert.equal(matchesNumericFilter(6, { max: 5 }), false);
  assert.equal(matchesNumericFilter(5, { min: 1, max: 10 }), true);
});

test("toCsv: plain cells join with commas and rows with CRLF", () => {
  const csv = toCsv(
    ["Part number", "Qty"],
    [
      ["WR17X11705", "5"],
      ["W11688994", "1"],
    ],
  );
  assert.equal(csv, "Part number,Qty\r\nWR17X11705,5\r\nW11688994,1");
});

test("toCsv: a cell with a comma, quote, or newline is quoted and internal quotes doubled", () => {
  const csv = toCsv(
    ["Description"],
    [['Say "hi", then stop'], ["multi\nline"]],
  );
  assert.equal(csv, 'Description\r\n"Say ""hi"", then stop"\r\n"multi\nline"');
});
