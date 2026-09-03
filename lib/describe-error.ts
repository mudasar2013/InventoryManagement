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
