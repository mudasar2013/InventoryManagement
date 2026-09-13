import assert from "node:assert/strict";
import { test } from "node:test";
import { parsePartInput, parsePartUpdateInput } from "./parse-part-input";

const validAddBody = {
  sourceId: "sharepoint-abc123",
  part_number: "WR17X11705",
  description: "Water filter cartridge",
  bin_location: "A-12-04",
  quantity_on_hand: 5,
};

test("parsePartInput: accepts a well-formed body", () => {
  const result = parsePartInput(validAddBody);
  assert.ok("input" in result);
  if ("input" in result) {
    assert.deepEqual(result.input, {
      sourceId: "sharepoint-abc123",
      fields: {
        part_number: "WR17X11705",
        description: "Water filter cartridge",
        bin_location: "A-12-04",
        quantity_on_hand: 5,
        category: "",
        extraFields: {},
      },
    });
  }
});

test("parsePartInput: carries category and extraFields through, re-validating each extra field's shape", () => {
  const result = parsePartInput({
    ...validAddBody,
    category: " Motor ",
    extraFields: {
      "Entry Date": { kind: "date", value: "2026-08-13" },
      "Ebay Ready (Yes/No)": { kind: "boolean", value: true },
      Condition: { kind: "select", value: "Used But Working", options: ["New", "Used", "Other"] },
      "HCPJob#": { kind: "text", value: "1234" },
      Garbage: { notAField: true },
    },
  });
  assert.ok("input" in result);
  if ("input" in result) {
    assert.equal(result.input.fields.category, "Motor");
    assert.deepEqual(result.input.fields.extraFields, {
      "Entry Date": { kind: "date", value: "2026-08-13" },
      "Ebay Ready (Yes/No)": { kind: "boolean", value: true },
      Condition: { kind: "select", value: "Used But Working", options: ["New", "Used", "Other"] },
      "HCPJob#": { kind: "text", value: "1234" },
    });
  }
});

test("parsePartInput: defaults description and bin_location to empty strings when omitted", () => {
  const result = parsePartInput({
    sourceId: "sharepoint-abc123",
    part_number: "WR17X11705",
    quantity_on_hand: 5,
  });
  assert.ok("input" in result);
  if ("input" in result) {
    assert.equal(result.input.fields.description, "");
    assert.equal(result.input.fields.bin_location, "");
  }
});

test("parsePartInput: rejects a missing or blank sourceId", () => {
  const result = parsePartInput({ ...validAddBody, sourceId: "  " });
  assert.ok("error" in result);
  if ("error" in result) {
    assert.match(result.error, /sourceId/);
  }
});

test("parsePartInput: rejects a missing or blank part_number", () => {
  const result = parsePartInput({ ...validAddBody, part_number: "" });
  assert.ok("error" in result);
  if ("error" in result) {
    assert.match(result.error, /part_number/);
  }
});

test("parsePartInput: rejects a non-numeric quantity_on_hand", () => {
  const result = parsePartInput({ ...validAddBody, quantity_on_hand: "a bunch" });
  assert.ok("error" in result);
  if ("error" in result) {
    assert.match(result.error, /quantity_on_hand/);
  }
});

test("parsePartInput: clamps a negative quantity_on_hand to 0", () => {
  const result = parsePartInput({ ...validAddBody, quantity_on_hand: -5 });
  assert.ok("input" in result);
  if ("input" in result) {
    assert.equal(result.input.fields.quantity_on_hand, 0);
  }
});

test("parsePartInput: rejects a non-object body", () => {
  assert.ok("error" in parsePartInput(null));
  assert.ok("error" in parsePartInput("a string"));
  assert.ok("error" in parsePartInput(undefined));
});

test("parsePartUpdateInput: accepts a well-formed body including existingPartNumber", () => {
  const result = parsePartUpdateInput({
    ...validAddBody,
    existingPartNumber: "WR17X11705-OLD",
  });
  assert.ok("input" in result);
  if ("input" in result) {
    assert.equal(result.input.existingPartNumber, "WR17X11705-OLD");
    assert.equal(result.input.fields.part_number, "WR17X11705");
  }
});

test("parsePartUpdateInput: rejects a missing existingPartNumber", () => {
  const result = parsePartUpdateInput(validAddBody);
  assert.ok("error" in result);
  if ("error" in result) {
    assert.match(result.error, /existingPartNumber/);
  }
});

test("parsePartUpdateInput: still reports the same validation errors as parsePartInput", () => {
  const result = parsePartUpdateInput({ ...validAddBody, part_number: "" });
  assert.ok("error" in result);
  if ("error" in result) {
    assert.match(result.error, /part_number/);
  }
});
