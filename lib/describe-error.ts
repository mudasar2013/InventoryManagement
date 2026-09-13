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

/**
 * The fuller counterpart to describeError() — for the "Data sources"
 * page's per-source debug view, not the shop-floor warning banner. That
 * page is used by whoever manages the SharePoint links (typos in a
 * hostname/path are the most common failure), so it's worth showing
 * everything that might pin down the cause: the full error name/message,
 * every nested .cause (Node's fetch/undici can nest network failures
 * several levels deep), and — for Microsoft Graph SDK errors — the
 * status code, error code, request id, and raw response body, which
 * carry the actually diagnostic detail when .message comes back empty.
 * Never includes the access token or any request header.
 */
export function describeErrorDetail(error: unknown): string {
  const lines: string[] = [];

  if (error instanceof Error) {
    lines.push(`${error.name}: ${error.message || "(empty message)"}`);
    appendGraphExtras(lines, error);

    let cause = (error as { cause?: unknown }).cause;
    let depth = 0;
    while (cause !== undefined && cause !== null && depth < 5) {
      if (cause instanceof Error) {
        lines.push(`Caused by: ${cause.name}: ${cause.message || "(empty message)"}`);
        // Steps get wrapped in a plain Error("<what was being attempted>",
        // { cause }) — see sharepoint-excel-source.ts's runStep() — so the
        // Graph SDK's own extras usually live on this .cause, not on the
        // outer wrapper. Check every node in the chain, not just the top.
        appendGraphExtras(lines, cause);
        cause = (cause as { cause?: unknown }).cause;
      } else {
        lines.push(`Caused by: ${String(cause)}`);
        cause = undefined;
      }
      depth += 1;
    }
  } else if (error !== null && error !== undefined) {
    lines.push(String(error));
  } else {
    lines.push("unknown error");
  }

  return lines.join("\n");
}

/** Appends Microsoft Graph SDK-shaped extras (status code, error code,
 *  request id, raw response body) found directly on one error object —
 *  the detail that actually pins down a cause when .message is generic
 *  or empty. Never includes the access token or any request header. */
function appendGraphExtras(lines: string[], error: Error): void {
  const withDetails = error as unknown as {
    statusCode?: number;
    code?: string | null;
    requestId?: string | null;
    body?: unknown;
  };

  if (typeof withDetails.statusCode === "number" && withDetails.statusCode !== -1) {
    lines.push(`HTTP status: ${withDetails.statusCode}`);
  }
  if (withDetails.code) {
    lines.push(`Code: ${withDetails.code}`);
  }
  if (withDetails.requestId) {
    lines.push(`Request id: ${withDetails.requestId}`);
  }
  if (withDetails.body !== undefined && withDetails.body !== null) {
    const bodyText =
      typeof withDetails.body === "string" ? withDetails.body : JSON.stringify(withDetails.body);
    if (bodyText && bodyText !== "{}" && bodyText !== "null") {
      lines.push(`Body: ${bodyText.slice(0, 500)}`);
    }
  }
}
