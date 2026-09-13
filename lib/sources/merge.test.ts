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
  const merged = mergeParts(
    [{ sourceId: "local", parts: [raw()] }],
    ["local"],
  );

  assert.equal(merged.length, 1);
  assert.equal(merged[0].part_number, "WR17X11705");
  assert.equal(merged[0].status, "In Stock");
  assert.deepEqual(merged[0].sourceIds, ["local"]);
});

test("mergeParts: the same part_number from two sources is deduped, not duplicated", () => {
  const merged = mergeParts(
    [
      { sourceId: "local", parts: [raw({ quantity_on_hand: 14 })] },
      { sourceId: "sharepoint", parts: [raw({ id: "sp-row-9", quantity_on_hand: 2 })] },
    ],
    ["local", "sharepoint"],
  );

  assert.equal(merged.length, 1);
});

test("mergeParts: conflicting fields resolve using the priority order, not arrival order", () => {
  const merged = mergeParts(
    [
      { sourceId: "sharepoint", parts: [raw({ id: "sp-row-9", quantity_on_hand: 2 })] },
      { sourceId: "local", parts: [raw({ quantity_on_hand: 14 })] },
    ],
    // "local" is ranked ahead of "sharepoint" even though sharepoint's
    // result was passed first — priority order must win, not call order.
    ["local", "sharepoint"],
  );

  assert.equal(merged[0].quantity_on_hand, 14);
  assert.equal(merged[0].id, "prt-local-1");
});

test("mergeParts: every reporting source is recorded even when one wins the conflict", () => {
  const merged = mergeParts(
    [
      { sourceId: "local", parts: [raw({ quantity_on_hand: 14 })] },
      { sourceId: "sharepoint", parts: [raw({ id: "sp-row-9", quantity_on_hand: 2 })] },
    ],
    ["local", "sharepoint"],
  );

  assert.deepEqual(new Set(merged[0].sourceIds), new Set(["local", "sharepoint"]));
});

test("mergeParts: the same source reporting the same part_number twice keeps both rows as separate parts", () => {
  // Mirrors a real sheet: an old row for a part number that previously
  // sold through (left in place, marked out of inventory) plus a freshly
  // appended row from restocking it. These are two genuinely different
  // physical rows, not one part described twice — a technician needs to
  // find and edit either one, so neither should shadow the other (see
  // RawPart.partNumberOccurrence, which is what keeps their ids distinct
  // in real sheet data).
  const merged = mergeParts(
    [
      {
        sourceId: "sharepoint",
        parts: [
          raw({ id: "sp-row-old", quantity_on_hand: 1, bin_location: "OLD-BIN" }),
          raw({ id: "sp-row-new", quantity_on_hand: 1, bin_location: "NEW-BIN" }),
        ],
      },
    ],
    ["sharepoint"],
  );

  assert.equal(merged.length, 2);
  const byId = new Map(merged.map((part) => [part.id, part]));
  assert.equal(byId.get("sp-row-old")?.bin_location, "OLD-BIN");
  assert.equal(byId.get("sp-row-new")?.bin_location, "NEW-BIN");
  assert.deepEqual(byId.get("sp-row-old")?.sourceIds, ["sharepoint"]);
  assert.deepEqual(byId.get("sp-row-new")?.sourceIds, ["sharepoint"]);
});

test("mergeParts: a same-source duplicate part_number doesn't stop the *first* occurrence from still merging with another source", () => {
  const merged = mergeParts(
    [
      {
        sourceId: "sharepoint",
        parts: [
          raw({ id: "sp-row-old", quantity_on_hand: 1, bin_location: "OLD-BIN" }),
          raw({ id: "sp-row-new", quantity_on_hand: 3, bin_location: "NEW-BIN" }),
        ],
      },
      { sourceId: "local", parts: [raw({ id: "local-row", quantity_on_hand: 99 })] },
    ],
    ["local", "sharepoint"],
  );

  assert.equal(merged.length, 2);
  const byId = new Map(merged.map((part) => [part.id, part]));
  // "local" outranks "sharepoint", so it wins the conflict over the
  // *first* sharepoint row -- the second (standalone) row is untouched.
  assert.equal(byId.get("local-row")?.quantity_on_hand, 99);
  assert.deepEqual(new Set(byId.get("local-row")?.sourceIds), new Set(["local", "sharepoint"]));
  assert.equal(byId.get("sp-row-new")?.quantity_on_hand, 3);
  assert.deepEqual(byId.get("sp-row-new")?.sourceIds, ["sharepoint"]);
});

test("mergeParts: a source absent from the priority list still merges, ranked last", () => {
  const merged = mergeParts(
    [
      { sourceId: "unranked", parts: [raw({ quantity_on_hand: 1 })] },
      { sourceId: "local", parts: [raw({ quantity_on_hand: 14 })] },
    ],
    ["local"],
  );

  assert.equal(merged[0].quantity_on_hand, 14);
});
