import {
  AdapterRateLimitError,
  AuthenticationError,
  NetworkError,
  PermissionError,
  ResourceNotFoundError,
} from "@chat-adapter/shared";

const ADAPTER_NAME = "outlook-email";

/**
 * Map Microsoft Graph HTTP errors to Chat SDK shared error types.
 * Follows the same pattern as the Teams adapter's error handler.
 */
export function handleGraphError(error: unknown, operation: string): never {
  if (error && typeof error === "object") {
    const err = error as Record<string, unknown>;
    const statusCode =
      (err.statusCode as number) ||
      (err.status as number) ||
      (err.code as number);

    if (statusCode === 401) {
      throw new AuthenticationError(
        ADAPTER_NAME,
        `Authentication failed for ${operation}: ${err.message || "unauthorized"}`
      );
    }

    if (statusCode === 403) {
      throw new PermissionError(ADAPTER_NAME, operation);
    }

    if (statusCode === 404) {
      throw new ResourceNotFoundError(
        ADAPTER_NAME,
        "message",
        typeof err.resourceId === "string" ? err.resourceId : undefined
      );
    }

    if (statusCode === 429) {
      const retryAfter =
        typeof err.retryAfter === "number" ? err.retryAfter : undefined;
      throw new AdapterRateLimitError(ADAPTER_NAME, retryAfter);
    }

    if (err.message && typeof err.message === "string") {
      throw new NetworkError(
        ADAPTER_NAME,
        `Graph API error during ${operation}: ${err.message}`,
        error instanceof Error ? error : undefined
      );
    }
  }

  throw new NetworkError(
    ADAPTER_NAME,
    `Graph API error during ${operation}: ${String(error)}`,
    error instanceof Error ? error : undefined
  );
}

/**
 * Error thrown when an adapter method is not supported by email.
 */
export class NotImplementedError extends Error {
  constructor(method: string) {
    super(`${method} is not supported by the Outlook email adapter`);
    this.name = "NotImplementedError";
  }
}
