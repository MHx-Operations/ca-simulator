// lib/tenants.ts — Load and validate the curated tenant list from public/tenants.json.
//
// Every function here is pure except loadTenants(), which wraps the one fetch
// call this module makes. No UI framework, no auth library: it only produces
// a validated TenantEntry[] for later phases to resolve against.

import { isGuid } from './guid';

export interface TenantEntry {
  name: string;
  tenantId: string;
}

/**
 * Compares two tenant ids the way MSAL's own `tid` claim and every
 * `account.tenantId` value in this codebase must be compared: MSAL's tenant
 * data is always lowercase, but `tenants.json` is operator-curated free text
 * and a mixed-case GUID there must still match. Use this at every point
 * where a stored/persisted/JSON-sourced tenant id meets an MSAL-sourced one.
 */
export function sameTenant(a: string | null | undefined, b: string | null | undefined): boolean {
  return !!a && !!b && a.toLowerCase() === b.toLowerCase();
}

/**
 * Validates and normalizes an arbitrary parsed JSON value into a tenant list.
 * Anything that does not fit the shape is dropped rather than surfaced — the
 * caller never sees a reason, only a smaller (or empty) list.
 */
export function parseTenantList(raw: unknown): TenantEntry[] {
  if (!Array.isArray(raw)) return [];

  // Both keys are deduplicated first-wins. Names are compared the same way
  // resolveTenant matches them (trimmed, lowercased): two entries that would
  // both answer `?tenant=<name>` are an ambiguity the parser resolves here,
  // rather than letting the resolver silently pick one.
  const seenIds = new Set<string>();
  const seenNames = new Set<string>();
  const result: TenantEntry[] = [];

  for (const entry of raw) {
    if (typeof entry !== 'object' || entry === null) continue;
    const { name, tenantId } = entry as Record<string, unknown>;
    if (typeof name !== 'string' || name.trim() === '') continue;
    if (typeof tenantId !== 'string' || !isGuid(tenantId)) continue;

    const idKey = tenantId.toLowerCase();
    const nameKey = name.trim().toLowerCase();
    if (seenIds.has(idKey) || seenNames.has(nameKey)) continue;
    seenIds.add(idKey);
    seenNames.add(nameKey);

    // Canonicalize to lowercase once, here at the validation boundary: MSAL
    // populates account.tenantId from the token's `tid` claim, which is
    // always lowercase, and its tenant-profile cache lookup is an exact-key
    // match. A mixed-case id from tenants.json would otherwise pass
    // validation and then silently fail every downstream comparison.
    result.push({ name: name.trim(), tenantId: idKey });
  }

  return result;
}

/**
 * Resolves a `?tenant=` value (or a persisted fallback) against the already-
 * validated tenant list. Matching is intentionally strict about what counts
 * as a match: the id side is compared case-insensitively (GUIDs have no
 * meaningful whitespace), the name side is trimmed and lowercased on the
 * incoming needle before comparison.
 *
 * The invariant: an unmatched value — including a syntactically valid
 * GUID that simply isn't in the list — is indistinguishable from an absent
 * one and falls through to the next candidate, eventually to `null`. No path
 * here ever synthesizes a `TenantEntry`; the alternative — trusting a string
 * that merely looks like a GUID — is how a crafted link turns into an
 * authority for a tenant the operator never curated.
 */
export function resolveTenant(
  tenants: TenantEntry[],
  params: { urlParam?: string | null; persistedId?: string | null },
): TenantEntry | null {
  // Ids are matched in a first pass, names in a second, so an entry whose
  // name happens to equal a later entry's id can never shadow the id match.
  const match = (needle: string): TenantEntry | undefined => {
    const n = needle.trim().toLowerCase();
    return tenants.find((t) => t.tenantId.toLowerCase() === n) ?? tenants.find((t) => t.name.trim().toLowerCase() === n);
  };

  if (params.urlParam) {
    const found = match(params.urlParam);
    if (found) return found;
  }
  if (params.persistedId) {
    const found = match(params.persistedId);
    if (found) return found;
  }
  return null;
}

/**
 * Rewrites a `location.search` string so its `tenant` parameter is the
 * canonical id, preserving every other query parameter. Pure: no DOM access.
 * `URLSearchParams` tolerates a leading `?` and handles encoding, so no
 * manual string surgery is needed.
 */
export function rewriteTenantParam(search: string, canonicalId: string): string {
  const params = new URLSearchParams(search);
  params.set('tenant', canonicalId);
  return `?${params.toString()}`;
}

/**
 * Applies rewriteTenantParam's result to the address bar via the History
 * API's replaceState call. The module's second and last impure function.
 *
 * Called from both boot legs in App.tsx, and only when the user actually
 * supplied a ?tenant= parameter — a tenant that resolved from the persisted
 * browser preference alone stays out of the address bar. The value written is
 * always the canonical id taken from the validated list (see resolveTenant),
 * never the raw incoming parameter. Writing the address bar is all it does:
 * adopting a deep-linked tenant as the remembered choice happens at sign-in
 * (useAuthStore's login()), not here and not at boot.
 */
export function applyCanonicalTenantParam(canonicalId: string): void {
  if (typeof window === 'undefined' || typeof history === 'undefined') return;
  try {
    // Keep the fragment: MSAL's redirect response arrives in `#code=...` and
    // is read asynchronously after boot. A query-only relative URL would
    // resolve to an empty fragment and erase it. `history.state` is passed
    // through for the same reason — this call rewrites one parameter, nothing else.
    const { search, hash } = window.location;
    history.replaceState(history.state, '', rewriteTenantParam(search, canonicalId) + hash);
  } catch {
    // History API blocked or unavailable — degrade silently.
  }
}

/**
 * Fetches /tenants.json and returns the validated tenant list, or [] for
 * every failure shape: absent file, network failure, empty body, an HTML
 * SPA-rewrite body, or malformed/wrong-shape JSON. Never throws, never logs.
 */
export async function loadTenants(fetchImpl: typeof fetch = fetch): Promise<TenantEntry[]> {
  try {
    const res = await fetchImpl('/tenants.json', { cache: 'no-store' });
    if (!res.ok) return [];
    const data = await res.json();
    return parseTenantList(data);
  } catch {
    return [];
  }
}

// --- Persistence -------------------------------------------------------------

/**
 * The app's third `ca-sim:` key, after `ca-sim:tour-seen` (lib/tour.ts) and
 * `ca-sim:analytics-opt-out` (lib/analytics.ts). Holds the canonical
 * `tenantId` of the last chosen tenant and nothing else — never the name,
 * because a name is display text that can be re-edited in tenants.json while
 * the id is the thing a later authority is built from. It is a preference
 * about this browser, and it is validated on read by list membership (see
 * resolveTenant) rather than trusted outright.
 */
export const TENANT_CHOICE_KEY = 'ca-sim:tenant';

/**
 * Returns the raw stored tenant id, or `null` if nothing is stored, storage
 * is blocked, or storage does not exist. Performs no GUID check by design —
 * validation happens by list membership in resolveTenant.
 */
export function readPersistedTenant(): string | null {
  if (typeof localStorage === 'undefined') return null;
  try {
    return localStorage.getItem(TENANT_CHOICE_KEY);
  } catch {
    return null;
  }
}

export function persistTenant(tenantId: string): void {
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(TENANT_CHOICE_KEY, tenantId);
  } catch {
    // Storage blocked — the choice simply will not survive this browser.
  }
}

// --- Authority -----------------------------------------------------------

/**
 * The upstream default authority baked into src/authConfig.ts's
 * buildMsalConfig(), duplicated here so this module can answer "what
 * authority with no tenant chosen?" without importing authConfig.ts (which
 * reads import.meta.env and window.location.origin inside that factory and
 * therefore cannot be imported into a node-environment unit test). A pinned
 * test guards the two copies against drift.
 */
export const DEFAULT_AUTHORITY = 'https://login.microsoftonline.com/organizations';

/**
 * Returns DEFAULT_AUTHORITY for no tenant. Entries normally come out of
 * parseTenantList/resolveTenant and are therefore GUID-validated, but
 * TenantEntry is a structural type and nothing stops a caller from building
 * one by hand. This is the single place a string becomes an MSAL authority,
 * so it re-checks the id itself: anything that is not a GUID falls back to
 * the default authority rather than being interpolated into a URL.
 */
export function authorityFor(tenant: TenantEntry | null): string {
  if (tenant === null || !isGuid(tenant.tenantId)) return DEFAULT_AUTHORITY;
  return `https://login.microsoftonline.com/${tenant.tenantId}`;
}
