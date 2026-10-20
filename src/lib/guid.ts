// lib/guid.ts — GUID shape check.
//
// Used for "is this an unresolvable object id?" display decisions and to
// validate hand-entered app IDs before they reach the engine or a Graph path.

export const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function isGuid(value: string): boolean {
  return GUID_RE.test(value);
}

/**
 * Replaces every GUID in a free-text line with `<id>`. For MSAL log lines that
 * reach the console: even lines MSAL does not flag as PII can quote an
 * authority URL, a tenant id or a correlation id.
 */
export function redactIds(message: string): string {
  return message.replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, '<id>');
}
