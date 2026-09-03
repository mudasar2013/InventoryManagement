/**
 * Turn a caught error into a short, human-readable phrase for a warning
 * banner. Plain Error objects have a useful .message, but the Microsoft
 * Graph SDK throws errors whose .message is empty and puts the actually
 * useful detail on .statusCode / .code / .body instead — found by
 * exercising the SharePoint source's failure path during testing.
 * Falls back gracefully so a warning is never a bare, unhelpful "()".
 */
export function describeError(error: unknown): string {
  if (error instanceof Error && error.message) {
    // Node's built-in fetch (undici) throws a generic "fetch failed" for
    // any network-layer problem (DNS, TLS/proxy interception, connection
    // refused, ...) and puts the actual reason on .cause instead of the
    // top-level message — surface it too, or "fetch failed" alone tells
    // nobody anything.
    const cause = (error as { cause?: unknown }).cause;
    if (cause instanceof Error && cause.message) {
      return `${error.message}: ${cause.message}`;
    }
    return error.message;
  }

  if (error && typeof error === "object") {
    const withDetails = error as {
      statusCode?: number;
      code?: string | null;
      body?: unknown;
    };

    const parts: string[] = [];
    if (withDetails.code) {
      parts.push(String(withDetails.code));
    }
    if (typeof withDetails.statusCode === "number") {
      parts.push(`HTTP ${withDetails.statusCode}`);
    }
    if (parts.length > 0) {
      return parts.join(", ");
    }
  }

  return "unknown error";
}
