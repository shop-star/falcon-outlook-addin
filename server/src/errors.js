/* One error shape for everything the task pane needs to explain to a
 * person: a stable code it can branch on, a sentence it can show as-is, and
 * (for rate limits) how long to wait before trying again.
 */
export class AppError extends Error {
  constructor(status, code, message, { retryAfter, details } = {}) {
    super(message);
    this.name = "AppError";
    this.status = status;
    this.code = code;
    this.retryAfter = retryAfter;
    this.details = details;
  }

  toJSON() {
    const body = { code: this.code, message: this.message };
    if (this.retryAfter !== undefined) body.retryAfter = this.retryAfter;
    if (this.details !== undefined) body.details = this.details;
    return { error: body };
  }
}

export function badRequest(message, details) {
  return new AppError(400, "BAD_REQUEST", message, { details });
}
