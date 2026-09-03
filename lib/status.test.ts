import assert from "node:assert/strict";
import { test } from "node:test";
import { deriveStatus, toPart } from "./status";

test("deriveStatus: zero or negative on hand is Out of Stock", () => {
  assert.equal(deriveStatus(0), "Out of Stock");
  assert.equal(deriveStatus(-1), "Out of Stock");
});

test("deriveStatus: at or under the threshold (but above zero) is Low Stock", () => {
  assert.equal(deriveStatus(1, 5), "Low Stock");
  assert.equal(deriveStatus(5, 5), "Low Stock");
});

test("deriveStatus: above the threshold is In Stock", () => {
  assert.equal(deriveStatus(6, 5), "In Stock");
});

test("deriveStatus: respects a custom threshold", () => {
  assert.equal(deriveStatus(2, 1), "In Stock");
  assert.equal(deriveStatus(1, 1), "Low Stock");
});

test("toPart: computes status instead of trusting a stored field", () => {
  const part = toPart({
    id: "prt-x",
    part_number: "X-1",
    description: "test part",
    bin_location: "Z-01",
    quantity_on_hand: 3,
  });
  assert.equal(part.status, "Low Stock");
});
