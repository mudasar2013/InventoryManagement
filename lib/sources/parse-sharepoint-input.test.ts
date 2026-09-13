import assert from "node:assert/strict";
import { test } from "node:test";
import { parseSharePointSourceInput } from "./parse-sharepoint-input";

const validBody = {
  label: "Downtown shop inventory",
  siteHostname: "contoso.sharepoint.com",
  sitePath: "/sites/ServiceOps",
  filePath: "Shared Documents/Inventory.xlsx",
  tableName: "Parts",
};

test("parseSharePointSourceInput: accepts a well-formed body as-is", () => {
  const result = parseSharePointSourceInput(validBody);
  assert.ok("input" in result);
  if ("input" in result) {
    assert.deepEqual(result.input, validBody);
  }
});

test("parseSharePointSourceInput: strips a pasted-in https:// scheme from siteHostname", () => {
  const result = parseSharePointSourceInput({
    ...validBody,
    siteHostname: "https://contoso.sharepoint.com",
  });
  assert.ok("input" in result);
  if ("input" in result) {
    assert.equal(result.input.siteHostname, "contoso.sharepoint.com");
  }
});

test("parseSharePointSourceInput: strips a pasted-in http:// scheme and trailing slash from siteHostname", () => {
  const result = parseSharePointSourceInput({
    ...validBody,
    siteHostname: "http://contoso.sharepoint.com/",
  });
  assert.ok("input" in result);
  if ("input" in result) {
    assert.equal(result.input.siteHostname, "contoso.sharepoint.com");
  }
});

test("parseSharePointSourceInput: adds a missing leading slash to sitePath", () => {
  const result = parseSharePointSourceInput({
    ...validBody,
    sitePath: "sites/ServiceOps",
  });
  assert.ok("input" in result);
  if ("input" in result) {
    assert.equal(result.input.sitePath, "/sites/ServiceOps");
  }
});

test("parseSharePointSourceInput: strips a trailing slash from sitePath", () => {
  const result = parseSharePointSourceInput({
    ...validBody,
    sitePath: "/sites/ServiceOps/",
  });
  assert.ok("input" in result);
  if ("input" in result) {
    assert.equal(result.input.sitePath, "/sites/ServiceOps");
  }
});

test("parseSharePointSourceInput: reports every missing or blank field", () => {
  const result = parseSharePointSourceInput({
    label: "",
    siteHostname: "contoso.sharepoint.com",
    sitePath: "  ",
  });
  assert.ok("error" in result);
  if ("error" in result) {
    assert.match(result.error, /label/);
    assert.match(result.error, /sitePath/);
    assert.match(result.error, /filePath/);
    assert.match(result.error, /tableName/);
  }
});

test("parseSharePointSourceInput: rejects a non-object body", () => {
  assert.ok("error" in parseSharePointSourceInput(null));
  assert.ok("error" in parseSharePointSourceInput("a string"));
  assert.ok("error" in parseSharePointSourceInput(undefined));
});
