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
