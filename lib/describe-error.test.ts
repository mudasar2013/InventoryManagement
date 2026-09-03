import assert from "node:assert/strict";
import { test } from "node:test";
import { describeError } from "./describe-error";

test("describeError: uses a plain Error's message", () => {
  assert.equal(describeError(new Error("network timeout")), "network timeout");
});

test("describeError: falls back to statusCode/code when message is empty (Graph SDK shape)", () => {
  const graphLikeError = Object.assign(new Error(""), {
    statusCode: 403,
    code: "Forbidden",
  });
  assert.equal(describeError(graphLikeError), "Forbidden, HTTP 403");
});

test("describeError: statusCode alone still produces something readable", () => {
  const error = Object.assign(new Error(""), { statusCode: 401, code: null });
  assert.equal(describeError(error), "HTTP 401");
});

test("describeError: never returns an empty or bare-punctuation string", () => {
  assert.equal(describeError(new Error("")), "unknown error");
  assert.equal(describeError("just a string"), "unknown error");
  assert.equal(describeError(null), "unknown error");
  assert.equal(describeError(undefined), "unknown error");
});
