import assert from "node:assert/strict";
import { test } from "node:test";
import { describeError, describeErrorDetail } from "./describe-error";

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

test("describeErrorDetail: includes name, message, and Graph SDK extras", () => {
  const graphLikeError = Object.assign(new Error("itemNotFound"), {
    statusCode: 404,
    code: "itemNotFound",
    requestId: "b31c83fd-944c-4663-aa50-5d9ceb367e19",
    body: JSON.stringify({ error: { code: "itemNotFound", message: "The resource could not be found." } }),
  });
  const detail = describeErrorDetail(graphLikeError);
  assert.match(detail, /^Error: itemNotFound$/m);
  assert.match(detail, /HTTP status: 404/);
  assert.match(detail, /Code: itemNotFound/);
  assert.match(detail, /Request id: b31c83fd-944c-4663-aa50-5d9ceb367e19/);
  assert.match(detail, /Body: .*The resource could not be found\./);
});

test("describeErrorDetail: walks nested .cause chains, not just one level", () => {
  const root = new Error("connect ECONNREFUSED 127.0.0.1:443");
  const middle = Object.assign(new Error("request to https://graph.microsoft.com failed"), {
    cause: root,
  });
  const top = Object.assign(new Error("fetch failed"), { cause: middle });
  const detail = describeErrorDetail(top);
  assert.match(detail, /^Error: fetch failed$/m);
  assert.match(detail, /Caused by: Error: request to https:\/\/graph\.microsoft\.com failed/);
  assert.match(detail, /Caused by: Error: connect ECONNREFUSED 127\.0\.0\.1:443/);
});

test("describeErrorDetail: reports an empty message rather than a blank line", () => {
  const graphLikeError = Object.assign(new Error(""), {
    statusCode: 403,
    code: "Forbidden",
  });
  assert.match(describeErrorDetail(graphLikeError), /\(empty message\)/);
});

test("describeErrorDetail: never returns an empty string", () => {
  assert.equal(describeErrorDetail(null), "unknown error");
  assert.equal(describeErrorDetail(undefined), "unknown error");
  assert.equal(describeErrorDetail("just a string"), "just a string");
});
