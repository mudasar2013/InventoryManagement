import assert from "node:assert/strict";
import { test } from "node:test";
import { upnFromExtraFields } from "./types";

test("upnFromExtraFields: reads a value under an exact \"UPN#\" header", () => {
  assert.equal(upnFromExtraFields({ "UPN#": { kind: "text", value: "596" } }), "596");
});

test("upnFromExtraFields: matches header spelling variants (case, spacing, missing #)", () => {
  assert.equal(upnFromExtraFields({ "upn #": { kind: "text", value: "1" } }), "1");
  assert.equal(upnFromExtraFields({ UPN: { kind: "text", value: "2" } }), "2");
  assert.equal(upnFromExtraFields({ " UPN# ": { kind: "text", value: "3" } }), "3");
});

test("upnFromExtraFields: returns undefined when there's no UPN-shaped column", () => {
  assert.equal(upnFromExtraFields({ Condition: { kind: "select", value: "New" } }), undefined);
  assert.equal(upnFromExtraFields(undefined), undefined);
  assert.equal(upnFromExtraFields({}), undefined);
});

test("upnFromExtraFields: returns undefined for a blank UPN# value rather than an empty string", () => {
  assert.equal(upnFromExtraFields({ "UPN#": { kind: "text", value: "" } }), undefined);
  assert.equal(upnFromExtraFields({ "UPN#": { kind: "text", value: "   " } }), undefined);
});
