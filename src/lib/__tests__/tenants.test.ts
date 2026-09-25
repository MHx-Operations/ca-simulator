// tenants.test.ts — Validation of the tenant list and the fetch/fallback loader.
//
// The module under test is deliberately React/MSAL-free: every failure shape
// (absent file, network error, HTML rewrite, malformed JSON) must degrade to
// an empty list without throwing or logging.

import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  parseTenantList,
  loadTenants,
  resolveTenant,
  rewriteTenantParam,
  applyCanonicalTenantParam,
  TENANT_CHOICE_KEY,
  readPersistedTenant,
  persistTenant,
  DEFAULT_AUTHORITY,
  authorityFor,
  sameTenant,
  type TenantEntry,
} from '../tenants';

const GOOD_A: TenantEntry = { name: 'Example Tenant', tenantId: '00000000-0000-0000-0000-000000000001' };
const GOOD_B: TenantEntry = { name: 'Second Example Tenant', tenantId: '00000000-0000-0000-0000-000000000002' };

describe('parseTenantList', () => {
  it('returns both entries in order, with name trimmed', () => {
    const raw = [
      { name: '  Example Tenant  ', tenantId: GOOD_A.tenantId },
      { name: 'Second Example Tenant', tenantId: GOOD_B.tenantId },
    ];
    expect(parseTenantList(raw)).toEqual([GOOD_A, GOOD_B]);
  });

  it.each([null, undefined, {}, 'text', 42])('returns [] for non-array input %p', (input) => {
    expect(parseTenantList(input)).toEqual([]);
  });

  it('returns [] for an empty array', () => {
    expect(parseTenantList([])).toEqual([]);
  });

  it.each(['not-a-guid', '', '{00000000-0000-0000-0000-000000000001}'])(
    'drops an entry with a non-GUID tenantId (%p) while a valid sibling survives',
    (badId) => {
      const raw = [{ name: 'Bad', tenantId: badId }, { name: GOOD_A.name, tenantId: GOOD_A.tenantId }];
      expect(parseTenantList(raw)).toEqual([GOOD_A]);
    },
  );

  it.each([undefined, 5, '   '])(
    'drops an entry with a missing/non-string/whitespace-only name (%p) while a valid sibling survives',
    (badName) => {
      const raw = [{ name: badName, tenantId: GOOD_B.tenantId }, { name: GOOD_A.name, tenantId: GOOD_A.tenantId }];
      expect(parseTenantList(raw)).toEqual([GOOD_A]);
    },
  );

  it.each([null, 'x', 5])('drops a non-object entry (%p) without throwing', (entry) => {
    const raw = [entry, { name: GOOD_A.name, tenantId: GOOD_A.tenantId }];
    expect(() => parseTenantList(raw)).not.toThrow();
    expect(parseTenantList(raw)).toEqual([GOOD_A]);
  });

  it('dedupes entries whose tenantId differs only in letter case, first occurrence wins', () => {
    const raw = [
      { name: 'First', tenantId: GOOD_A.tenantId.toUpperCase() },
      { name: 'Second', tenantId: GOOD_A.tenantId },
    ];
    expect(parseTenantList(raw)).toEqual([{ name: 'First', tenantId: GOOD_A.tenantId }]);
  });

  it('canonicalizes a mixed-case tenantId to lowercase', () => {
    const mixedCase = 'ABCDEF01-1234-5678-9ABC-DEF012345678';
    const raw = [{ name: 'Mixed Case Tenant', tenantId: mixedCase }];
    expect(parseTenantList(raw)[0].tenantId).toBe(mixedCase.toLowerCase());
  });

  it('dedupes entries whose name differs only in case/whitespace, first occurrence wins', () => {
    const raw = [
      { name: 'Example Tenant', tenantId: GOOD_A.tenantId },
      { name: '  example tenant ', tenantId: GOOD_B.tenantId },
    ];
    expect(parseTenantList(raw)).toEqual([GOOD_A]);
  });
});

describe('loadTenants', () => {
  it('returns the parsed entry on the happy path and fetches /tenants.json', async () => {
    const mockFetch = vi.fn().mockResolvedValue(
      new Response(JSON.stringify([{ name: 'Example Tenant', tenantId: GOOD_A.tenantId }]), { status: 200 }),
    );
    expect(await loadTenants(mockFetch)).toEqual([GOOD_A]);
    expect(mockFetch).toHaveBeenCalledWith('/tenants.json', expect.anything());
  });

  it('returns [] on a 404', async () => {
    const mockFetch = vi.fn().mockResolvedValue(new Response(null, { status: 404 }));
    expect(await loadTenants(mockFetch)).toEqual([]);
  });

  it('returns [] when fetch rejects (network failure) and does not throw', async () => {
    const mockFetch = vi.fn().mockRejectedValue(new Error('offline'));
    await expect(loadTenants(mockFetch)).resolves.toEqual([]);
  });

  it('returns [] when fetch throws synchronously and does not throw', async () => {
    const mockFetch = vi.fn(() => {
      throw new Error('blocked');
    });
    await expect(loadTenants(mockFetch)).resolves.toEqual([]);
  });

  it('returns [] for an HTML SPA-rewrite body', async () => {
    const mockFetch = vi.fn().mockResolvedValue(
      new Response('<!doctype html><html></html>', { status: 200, headers: { 'content-type': 'text/html' } }),
    );
    expect(await loadTenants(mockFetch)).toEqual([]);
  });

  it('returns [] for an empty body', async () => {
    const mockFetch = vi.fn().mockResolvedValue(new Response('', { status: 200 }));
    expect(await loadTenants(mockFetch)).toEqual([]);
  });

  it('returns [] for valid-JSON-but-wrong-shape body', async () => {
    const mockFetch = vi.fn().mockResolvedValue(new Response('{"tenants":[]}', { status: 200 }));
    expect(await loadTenants(mockFetch)).toEqual([]);
  });

  it('never logs to the console across any failure case', async () => {
    const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});

    await loadTenants(vi.fn().mockResolvedValue(new Response(null, { status: 404 })));
    await loadTenants(vi.fn().mockRejectedValue(new Error('offline')));
    await loadTenants(
      vi.fn(() => {
        throw new Error('blocked');
      }),
    );
    await loadTenants(
      vi.fn().mockResolvedValue(
        new Response('<!doctype html><html></html>', { status: 200, headers: { 'content-type': 'text/html' } }),
      ),
    );
    await loadTenants(vi.fn().mockResolvedValue(new Response('', { status: 200 })));
    await loadTenants(vi.fn().mockResolvedValue(new Response('{"tenants":[]}', { status: 200 })));

    expect(errorSpy).not.toHaveBeenCalled();
    expect(warnSpy).not.toHaveBeenCalled();

    errorSpy.mockRestore();
    warnSpy.mockRestore();
  });
});

describe('resolveTenant', () => {
  const LIST: TenantEntry[] = [GOOD_A, GOOD_B];

  it('returns the entry matching a urlParam id', () => {
    expect(resolveTenant(LIST, { urlParam: GOOD_B.tenantId })).toBe(GOOD_B);
  });

  it('matches the urlParam id case-insensitively', () => {
    expect(resolveTenant(LIST, { urlParam: GOOD_B.tenantId.toUpperCase() })).toBe(GOOD_B);
  });

  it('matches the urlParam name case-insensitively', () => {
    expect(resolveTenant(LIST, { urlParam: 'example tenant' })).toBe(GOOD_A);
  });

  it('trims surrounding whitespace on the urlParam name', () => {
    expect(resolveTenant(LIST, { urlParam: '  Example Tenant  ' })).toBe(GOOD_A);
  });

  it('returns null for an unmatched urlParam when nothing is persisted', () => {
    expect(resolveTenant(LIST, { urlParam: 'Nonexistent' })).toBeNull();
  });

  it('returns null for a syntactically valid but unlisted GUID, never a synthesized entry', () => {
    const result = resolveTenant(LIST, { urlParam: '99999999-9999-9999-9999-999999999999' });
    expect(result).toBeNull();
  });

  it.each([{}, { urlParam: null }, { urlParam: '' }])('returns null for %p', (params) => {
    expect(resolveTenant(LIST, params)).toBeNull();
  });

  it('returns null against an empty list regardless of urlParam', () => {
    expect(resolveTenant([], { urlParam: GOOD_A.tenantId })).toBeNull();
  });

  it('prefers urlParam over persistedId when both match', () => {
    expect(resolveTenant(LIST, { urlParam: GOOD_B.tenantId, persistedId: GOOD_A.tenantId })).toBe(GOOD_B);
  });

  it('falls through to persistedId when urlParam does not match', () => {
    expect(resolveTenant(LIST, { urlParam: 'Nonexistent', persistedId: GOOD_A.tenantId })).toBe(GOOD_A);
  });

  it('returns null for a garbage persistedId', () => {
    expect(resolveTenant(LIST, { persistedId: 'garbage-value' })).toBeNull();
  });

  it('prefers an id match over an earlier entry whose name equals that id', () => {
    const shadow: TenantEntry = { name: GOOD_B.tenantId, tenantId: GOOD_A.tenantId };
    expect(resolveTenant([shadow, GOOD_B], { urlParam: GOOD_B.tenantId })).toBe(GOOD_B);
  });
});

// ── Persistence ──

function stubStorage(): Map<string, string> {
  const store = new Map<string, string>();
  vi.stubGlobal('localStorage', {
    getItem: (k: string) => store.get(k) ?? null,
    setItem: (k: string, v: string) => void store.set(k, v),
    removeItem: (k: string) => void store.delete(k),
  });
  return store;
}

describe('persistTenant / readPersistedTenant', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('is null on first visit and returns the stored value once persisted', () => {
    const store = stubStorage();
    expect(readPersistedTenant()).toBeNull();
    persistTenant(GOOD_A.tenantId);
    expect(store.get(TENANT_CHOICE_KEY)).toBe(GOOD_A.tenantId);
    expect(readPersistedTenant()).toBe(GOOD_A.tenantId);
  });

  it('overwrites a previously persisted id with a second write', () => {
    stubStorage();
    persistTenant(GOOD_A.tenantId);
    persistTenant(GOOD_B.tenantId);
    expect(readPersistedTenant()).toBe(GOOD_B.tenantId);
  });

  it('survives blocked storage rather than throwing', () => {
    const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});
    vi.stubGlobal('localStorage', {
      getItem: () => {
        throw new Error('access denied');
      },
      setItem: () => {
        throw new Error('access denied');
      },
      removeItem: () => {},
    });

    expect(() => persistTenant(GOOD_A.tenantId)).not.toThrow();
    expect(readPersistedTenant()).toBeNull();
    expect(errorSpy).not.toHaveBeenCalled();
    expect(warnSpy).not.toHaveBeenCalled();

    errorSpy.mockRestore();
    warnSpy.mockRestore();
  });

  it('returns null and does not throw when there is no storage at all', () => {
    const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});
    vi.stubGlobal('localStorage', undefined);

    expect(readPersistedTenant()).toBeNull();
    expect(() => persistTenant(GOOD_A.tenantId)).not.toThrow();
    expect(errorSpy).not.toHaveBeenCalled();
    expect(warnSpy).not.toHaveBeenCalled();

    errorSpy.mockRestore();
    warnSpy.mockRestore();
  });
});

describe('precedence through storage', () => {
  const LIST: TenantEntry[] = [GOOD_A, GOOD_B];

  afterEach(() => vi.unstubAllGlobals());

  it('preselects the persisted entry on the next visit', () => {
    stubStorage();
    persistTenant(GOOD_A.tenantId);
    expect(resolveTenant(LIST, { persistedId: readPersistedTenant() })).toBe(GOOD_A);
  });

  it('lets a urlParam override the persisted choice', () => {
    stubStorage();
    persistTenant(GOOD_A.tenantId);
    expect(resolveTenant(LIST, { urlParam: GOOD_B.tenantId, persistedId: readPersistedTenant() })).toBe(GOOD_B);
  });

  it('resolves to null when the persisted id is no longer in the list', () => {
    stubStorage();
    persistTenant(GOOD_A.tenantId);
    expect(resolveTenant([GOOD_B], { persistedId: readPersistedTenant() })).toBeNull();
  });

  it.each(['not-a-guid', '{}', ''])('resolves a corrupted persisted value (%p) to null without throwing', (bad) => {
    stubStorage();
    if (bad !== '') persistTenant(bad);
    expect(() => resolveTenant(LIST, { persistedId: readPersistedTenant() })).not.toThrow();
    expect(resolveTenant(LIST, { persistedId: readPersistedTenant() })).toBeNull();
  });

  // Mirrors the normal-boot leg in App.tsx (this project has no component
  // tests): the address bar is rewritten only for a ?tenant= the user
  // supplied, never to promote the private persisted choice into a URL.
  it('does not rewrite the address bar when only the persisted choice resolved and the URL has no tenant parameter', () => {
    stubStorage();
    persistTenant(GOOD_A.tenantId);
    const replaceState = vi.fn();
    const search = '?view=grid';
    vi.stubGlobal('window', { location: { search, hash: '' } });
    vi.stubGlobal('history', { replaceState, state: null });
    const params = new URLSearchParams(search);

    const resolved = resolveTenant(LIST, { urlParam: params.get('tenant'), persistedId: readPersistedTenant() });
    if (resolved && params.has('tenant')) applyCanonicalTenantParam(resolved.tenantId);

    expect(resolved).toBe(GOOD_A);
    expect(replaceState).not.toHaveBeenCalled();
  });

  it('rewrites the address bar when the user supplied a ?tenant= name that resolved', () => {
    stubStorage();
    persistTenant(GOOD_A.tenantId);
    const replaceState = vi.fn();
    const search = `?tenant=${encodeURIComponent(GOOD_B.name)}`;
    vi.stubGlobal('window', { location: { search, hash: '' } });
    vi.stubGlobal('history', { replaceState, state: null });
    const params = new URLSearchParams(search);

    const resolved = resolveTenant(LIST, { urlParam: params.get('tenant'), persistedId: readPersistedTenant() });
    if (resolved && params.has('tenant')) applyCanonicalTenantParam(resolved.tenantId);

    expect(resolved).toBe(GOOD_B);
    expect(replaceState).toHaveBeenCalledWith(null, '', rewriteTenantParam(search, GOOD_B.tenantId));
  });
});

describe('sameTenant', () => {
  it('matches ids that differ only in case', () => {
    expect(sameTenant(GOOD_A.tenantId.toUpperCase(), GOOD_A.tenantId)).toBe(true);
  });

  it('does not match different ids', () => {
    expect(sameTenant(GOOD_A.tenantId, GOOD_B.tenantId)).toBe(false);
  });

  it.each([null, undefined])('returns false when either side is %p', (nullish) => {
    expect(sameTenant(nullish, GOOD_A.tenantId)).toBe(false);
    expect(sameTenant(GOOD_A.tenantId, nullish)).toBe(false);
  });
});

describe('authorityFor / DEFAULT_AUTHORITY', () => {
  it('pins DEFAULT_AUTHORITY to the upstream organizations authority', () => {
    expect(DEFAULT_AUTHORITY).toBe('https://login.microsoftonline.com/organizations');
  });

  it('returns DEFAULT_AUTHORITY for no tenant', () => {
    expect(authorityFor(null)).toBe(DEFAULT_AUTHORITY);
  });

  it('returns the tenant-scoped authority for a resolved entry', () => {
    expect(authorityFor(GOOD_A)).toBe(`https://login.microsoftonline.com/${GOOD_A.tenantId}`);
  });

  it('never embeds the entry name in the authority', () => {
    const result = authorityFor(GOOD_A);
    expect(result).not.toContain(GOOD_A.name);
  });

  it.each(['common/../evil', 'not-a-guid', '', 'organizations'])(
    'returns DEFAULT_AUTHORITY for a hand-built entry whose tenantId is not a GUID (%p)',
    (badId) => {
      expect(authorityFor({ name: 'x', tenantId: badId })).toBe(DEFAULT_AUTHORITY);
    },
  );
});

describe('rewriteTenantParam', () => {
  const CANONICAL = '00000000-0000-0000-0000-000000000001';

  it('sets the tenant parameter to the canonical id', () => {
    const result = rewriteTenantParam('?tenant=Example%20Tenant', CANONICAL);
    expect(new URLSearchParams(result).get('tenant')).toBe(CANONICAL);
  });

  it('preserves other parameters', () => {
    const result = rewriteTenantParam('?view=matrix&tenant=Example', CANONICAL);
    const params = new URLSearchParams(result);
    expect(params.get('view')).toBe('matrix');
    expect(params.get('tenant')).toBe(CANONICAL);
  });

  it('adds a tenant key when none is present', () => {
    const result = rewriteTenantParam('?view=matrix', CANONICAL);
    expect(new URLSearchParams(result).get('tenant')).toBe(CANONICAL);
  });

  it('returns ?tenant=<id> for an empty search string', () => {
    expect(rewriteTenantParam('', CANONICAL)).toBe(`?tenant=${CANONICAL}`);
  });

  it('accepts input with or without a leading ?', () => {
    expect(rewriteTenantParam('?a=1', CANONICAL)).toBe(rewriteTenantParam('a=1', CANONICAL));
  });

  it('always returns a string starting with ? that round-trips to the canonical id', () => {
    const result = rewriteTenantParam('a=1&b=2', CANONICAL);
    expect(result.startsWith('?')).toBe(true);
    expect(new URLSearchParams(result).get('tenant')).toBe(CANONICAL);
  });
});

describe('applyCanonicalTenantParam', () => {
  const CANONICAL = '00000000-0000-0000-0000-000000000002';

  afterEach(() => vi.unstubAllGlobals());

  it('calls history.replaceState once with the rewritten search string', () => {
    const replaceState = vi.fn();
    vi.stubGlobal('window', { location: { search: '?view=grid', hash: '' } });
    vi.stubGlobal('history', { replaceState, state: null });

    applyCanonicalTenantParam(CANONICAL);

    expect(replaceState).toHaveBeenCalledTimes(1);
    expect(replaceState).toHaveBeenCalledWith(null, '', rewriteTenantParam('?view=grid', CANONICAL));
  });

  it('preserves the URL fragment so an MSAL #code= redirect response survives the rewrite', () => {
    const replaceState = vi.fn();
    vi.stubGlobal('window', { location: { search: '?tenant=Example', hash: '#code=abc' } });
    vi.stubGlobal('history', { replaceState, state: null });

    applyCanonicalTenantParam(CANONICAL);

    expect(replaceState).toHaveBeenCalledTimes(1);
    const url = replaceState.mock.calls[0][2] as string;
    expect(url.endsWith('#code=abc')).toBe(true);
    expect(new URLSearchParams(url.slice(0, url.indexOf('#'))).get('tenant')).toBe(CANONICAL);
  });

  it('passes the existing history.state through instead of replacing it with null', () => {
    const replaceState = vi.fn();
    const state = { from: 'test' };
    vi.stubGlobal('window', { location: { search: '', hash: '' } });
    vi.stubGlobal('history', { replaceState, state });

    applyCanonicalTenantParam(CANONICAL);

    expect(replaceState).toHaveBeenCalledWith(state, '', rewriteTenantParam('', CANONICAL));
  });

  it('is a no-op and does not throw when window is undefined', () => {
    vi.stubGlobal('window', undefined);
    expect(() => applyCanonicalTenantParam(CANONICAL)).not.toThrow();
  });

  it('does not throw in the default Vitest node environment with nothing stubbed', () => {
    expect(() => applyCanonicalTenantParam(CANONICAL)).not.toThrow();
  });

  it('does not throw and logs nothing when history.replaceState throws', () => {
    const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    vi.stubGlobal('window', { location: { search: '?view=grid', hash: '' } });
    vi.stubGlobal('history', {
      state: null,
      replaceState: () => {
        throw new Error('blocked');
      },
    });

    expect(() => applyCanonicalTenantParam(CANONICAL)).not.toThrow();
    expect(errorSpy).not.toHaveBeenCalled();

    errorSpy.mockRestore();
  });
});
