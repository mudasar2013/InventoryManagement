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

test("describeError: appends a nested .cause (undici's generic 'fetch failed' shape)", () => {
  const causeError = new Error("getaddrinfo ENOTFOUND graph.microsoft.com");
  const fetchFailed = Object.assign(new Error("fetch failed"), {
    cause: causeError,
  });
  assert.equal(
    describeError(fetchFailed),
    "fetch failed: getaddrinfo ENOTFOUND graph.microsoft.com",
  );
});

test("describeError: ignores a .cause that isn't an Error with a message", () => {
  const fetchFailed = Object.assign(new Error("fetch failed"), {
    cause: "not an error object",
  });
  assert.equal(describeError(fetchFailed), "fetch failed");
});

test("describeError: never returns an empty or bare-punctuation string", () => {
  assert.equal(describeError(new Error("")), "unknown error");
  assert.equal(describeError("just a string"), "unknown error");
  assert.equal(describeError(null), "unknown error");
  assert.equal(describeError(undefined), "unknown error");
});
