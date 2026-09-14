import assert from "node:assert/strict";
import { test } from "node:test";
import { distinctValues, linkPartToJob } from "./inventory";
import { jobParts, parts } from "./mockData";

const quantitiesBefore = parts.map((part) => ({
  id: part.id,
  quantity_on_hand: part.quantity_on_hand,
}));

const linked = linkPartToJob(jobParts, "job-1042", "prt-004");

assert.equal(linked.length, jobParts.length + 1);
assert.equal(
  linked.some((item) => item.job_id === "job-1042" && item.part_id === "prt-004"),
  true,
);
assert.deepEqual(
  parts.map((part) => ({ id: part.id, quantity_on_hand: part.quantity_on_hand })),
  quantitiesBefore,
);

const again = linkPartToJob(linked, "job-1042", "prt-004");
assert.equal(again.length, linked.length);
assert.deepEqual(
  parts.map((part) => ({ id: part.id, quantity_on_hand: part.quantity_on_hand })),
  quantitiesBefore,
);

console.log("linkPartToJob does not deduct inventory");

test("distinctValues: dedupes, trims, drops blanks, and sorts alphabetically", () => {
  assert.deepEqual(
    distinctValues(["B-1", "A-2", " B-1 ", "", undefined, "   ", "A-2"]),
    ["A-2", "B-1"],
  );
});

test("distinctValues: returns an empty array when nothing is left after trimming", () => {
  assert.deepEqual(distinctValues([undefined, "", "   "]), []);
});
