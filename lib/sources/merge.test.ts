import assert from "node:assert/strict";
import { test } from "node:test";
import { mergeParts } from "./merge";

function raw(overrides: Partial<Parameters<typeof mergeParts>[0][number]["parts"][number]> = {}) {
  return {
    id: "prt-local-1",
    part_number: "WR17X11705",
    description: "Refrigerator water filter cartridge",
    bin_location: "A-12-04",
    quantity_on_hand: 14,
    ...overrides,
  };
}

test("mergeParts: a part_number reported by only one source passes through, status derived", () => {
  const merged = mergeParts([{ sourceId: "local", parts: [raw()] }]);

  assert.equal(merged.length, 1);
  assert.equal(merged[0].part_number, "WR17X11705");
  assert.equal(merged[0].status, "In Stock");
  assert.deepEqual(merged[0].sourceIds, ["local"]);
});

test("mergeParts: the same part_number from two sources stays two distinct Parts, not deduped", () => {
  // "Hadi Inventory" and "7300 Inventory" reporting the same part_number
  // are two different physical stock records (different bin, different
  // quantity) — combining them into one card would silently drop one
  // source's data and make its row impossible to find or edit on its
  // own, so both must survive as separate Parts.
  const merged = mergeParts([
    { sourceId: "hadi-inventory", parts: [raw({ id: "hadi-w11688994", quantity_on_hand: 1, bin_location: "L02-R1-S3-B" })] },
    { sourceId: "7300-inventory", parts: [raw({ id: "7300-w11688994", quantity_on_hand: 1, bin_location: "R0-S4-B8-Clear" })] },
  ]);

  assert.equal(merged.length, 2);
  const byId = new Map(merged.map((part) => [part.id, part]));
  assert.equal(byId.get("hadi-w11688994")?.bin_location, "L02-R1-S3-B");
  assert.deepEqual(byId.get("hadi-w11688994")?.sourceIds, ["hadi-inventory"]);
  assert.equal(byId.get("7300-w11688994")?.bin_location, "R0-S4-B8-Clear");
  assert.deepEqual(byId.get("7300-w11688994")?.sourceIds, ["7300-inventory"]);
});

test("mergeParts: the same source reporting the same part_number twice also stays two distinct Parts", () => {
  // Mirrors a real sheet: an old row for a part number that previously
  // sold through (left in place, marked out of inventory) plus a freshly
  // appended row from restocking it. These are two genuinely different
  // physical rows, not one part described twice.
  const merged = mergeParts([
    {
      sourceId: "sharepoint",
      parts: [
        raw({ id: "sp-row-old", quantity_on_hand: 1, bin_location: "OLD-BIN" }),
        raw({ id: "sp-row-new", quantity_on_hand: 1, bin_location: "NEW-BIN" }),
      ],
    },
  ]);

  assert.equal(merged.length, 2);
  const byId = new Map(merged.map((part) => [part.id, part]));
  assert.equal(byId.get("sp-row-old")?.bin_location, "OLD-BIN");
  assert.equal(byId.get("sp-row-new")?.bin_location, "NEW-BIN");
  assert.deepEqual(byId.get("sp-row-old")?.sourceIds, ["sharepoint"]);
  assert.deepEqual(byId.get("sp-row-new")?.sourceIds, ["sharepoint"]);
});

test("mergeParts: an empty source contributes nothing", () => {
  const merged = mergeParts([
    { sourceId: "local", parts: [raw()] },
    { sourceId: "sharepoint", parts: [] },
  ]);

  assert.equal(merged.length, 1);
  assert.deepEqual(merged[0].sourceIds, ["local"]);
});

test("mergeParts: preserves each row's own field values rather than picking a winner", () => {
  const merged = mergeParts([
    { sourceId: "sharepoint", parts: [raw({ id: "sp-row-9", quantity_on_hand: 2 })] },
    { sourceId: "local", parts: [raw({ quantity_on_hand: 14 })] },
  ]);

  const byId = new Map(merged.map((part) => [part.id, part]));
  assert.equal(byId.get("sp-row-9")?.quantity_on_hand, 2);
  assert.equal(byId.get("prt-local-1")?.quantity_on_hand, 14);
});
