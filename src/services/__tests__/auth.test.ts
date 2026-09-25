// services/__tests__/auth.test.ts — Tenant-scoped token acquisition in
// getAccessToken: account resolution by homeAccountId, the per-request
// authority, the response-tenant assertion and the two interaction-required
// paths.
//
// Uses the real useAuthStore (importable in Vitest's node environment — no
// DOM needed for setState) with a minimal fake MSAL instance object, per
// this project's convention of mocking only at I/O boundaries rather than
// the store itself.

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import type { PublicClientApplication } from '@azure/msal-browser';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { getAccessToken, TENANT_MISMATCH_ERROR } from '../auth';
import { useAuthStore } from '../../stores/useAuthStore';
import type { TenantEntry } from '../../lib/tenants';

const TENANT_A: TenantEntry = { name: 'Tenant A', tenantId: '11111111-1111-1111-1111-111111111111' };
const TENANT_B: TenantEntry = { name: 'Tenant B', tenantId: '22222222-2222-2222-2222-222222222222' };

const HOME_ACCOUNT_ID = 'home.signed-in';

interface AccountFilter {
  homeAccountId?: string;
  tenantId?: string;
  isHomeTenant?: boolean;
}

/**
 * Models MSAL's ONE-PROFILE behaviour, as observed against live tenants:
 * a single account entity keyed by homeAccountId whose tenantId is the tenant
 * of the most recently minted token, and an active-account filter that keeps
 * naming the tenant profile it was set to — so getActiveAccount() answers null
 * once a token for another tenant has been issued. getAllAccounts is the only
 * lookup that survives, and only through its homeAccountId filter.
 */
function makeOneProfileInstance(
  options: {
    /** The tenant whose profile the account currently carries. */
    profileTenantId?: string;
    acquireTokenSilent?: (request?: unknown) => Promise<unknown>;
  } = {},
) {
  const profileTenantId = options.profileTenantId ?? TENANT_A.tenantId;
  const entity = { homeAccountId: HOME_ACCOUNT_ID, tenantId: profileTenantId, username: 'user@example.test' };
  const activeProfileTenantId = profileTenantId;

  return {
    getActiveAccount: vi.fn(() => (entity.tenantId === activeProfileTenantId ? { ...entity } : null)),
    getAllAccounts: vi.fn((filter?: AccountFilter) => {
      if (filter?.homeAccountId) return filter.homeAccountId === entity.homeAccountId ? [{ ...entity }] : [];
      if (filter?.tenantId) return filter.tenantId === entity.tenantId ? [{ ...entity }] : [];
      if (filter?.isHomeTenant) return [];
      return [{ ...entity }];
    }),
    setActiveAccount: vi.fn(),
    acquireTokenSilent: vi.fn(
      options.acquireTokenSilent ??
        (async (request?: unknown) => {
          // Minting a token for a realm replaces the account's only tenant
          // profile — the previous tenant's profile is dropped, not merged.
          const tenantId = (request as { account: { tenantId: string } }).account.tenantId;
          entity.tenantId = tenantId;
          return { accessToken: 'token-abc', tenantId };
        }),
    ),
    acquireTokenRedirect: vi.fn(async () => {}),
  };
}

/**
 * Two distinct cached USERS — the leftover-account case (an abandoned logout
 * leaves msal.* keys behind, so the next user signs in over them). Only the
 * homeAccountId filter distinguishes them; an unfiltered [0] does not.
 */
function makeTwoUserInstance(entities: Array<{ homeAccountId: string; tenantId: string; username: string }>) {
  return {
    // Mirrors the post-switch reality: the active-account filter names a
    // tenant profile that no longer exists.
    getActiveAccount: vi.fn(() => null),
    getAllAccounts: vi.fn((filter?: AccountFilter) => {
      const list = filter?.homeAccountId
        ? entities.filter((e) => e.homeAccountId === filter.homeAccountId)
        : entities;
      return list.map((e) => ({ ...e }));
    }),
    setActiveAccount: vi.fn(),
    acquireTokenSilent: vi.fn(async (request?: unknown) => ({
      accessToken: 'token-abc',
      tenantId: (request as { account: { tenantId: string } }).account.tenantId,
    })),
    acquireTokenRedirect: vi.fn(async () => {}),
  };
}

/** Minimal fake for the upstream (no tenant selected) path. */
function makeUpstreamInstance(overrides: {
  getActiveAccount?: () => unknown;
  getAllAccounts?: (filter?: unknown) => unknown[];
  acquireTokenSilent?: (request?: unknown) => Promise<unknown>;
} = {}) {
  return {
    getActiveAccount: vi.fn(overrides.getActiveAccount ?? (() => null)),
    getAllAccounts: vi.fn(overrides.getAllAccounts ?? (() => [])),
    setActiveAccount: vi.fn(),
    acquireTokenSilent: vi.fn(
      overrides.acquireTokenSilent ?? (async () => ({ accessToken: 'token-abc', tenantId: TENANT_A.tenantId })),
    ),
    acquireTokenRedirect: vi.fn(async () => {}),
  };
}

function makeAccount(tenantId: string) {
  return { homeAccountId: `home.${tenantId}`, tenantId, username: 'user@example.test' };
}

// redirectInProgress is a module-scope latch in services/auth.ts that never
// resets (the redirect unloads the page). Cases that need to observe a first
// redirect get a fresh module pair — the fresh auth module closes over a
// fresh useAuthStore, so state must be set on the store it returns.
async function importFreshAuth() {
  vi.resetModules();
  const [authMod, storeMod] = await Promise.all([
    import('../auth'),
    import('../../stores/useAuthStore'),
  ]);
  return { getAccessToken: authMod.getAccessToken, authStore: storeMod.useAuthStore };
}

beforeEach(() => {
  useAuthStore.setState({ msalInstance: null, selectedTenant: null, homeAccountId: null });
});

afterEach(() => {
  vi.clearAllMocks();
});

describe('getAccessToken — account resolution by homeAccountId', () => {
  it('resolves the base account through the homeAccountId filter and stamps the selected tenant on it', async () => {
    const fakeInstance = makeOneProfileInstance();
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_A,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    const token = await getAccessToken();

    expect(fakeInstance.getAllAccounts).toHaveBeenCalledWith({ homeAccountId: HOME_ACCOUNT_ID });
    const request = fakeInstance.acquireTokenSilent.mock.calls[0][0] as Record<string, unknown>;
    const account = request.account as { homeAccountId: string; tenantId: string };
    expect(account.homeAccountId).toBe(HOME_ACCOUNT_ID);
    expect(account.tenantId).toBe(TENANT_A.tenantId);
    expect(request.authority).toBe(`https://login.microsoftonline.com/${TENANT_A.tenantId}`);
    expect(request).not.toHaveProperty('cacheLookupPolicy');
    expect(token).toBe('token-abc');
  });

  it('does not depend on getActiveAccount — acquisition succeeds while it answers null', async () => {
    // The account's only tenant profile names the OTHER tenant, so MSAL's
    // active-account filter resolves to nothing.
    const fakeInstance = makeOneProfileInstance({ profileTenantId: TENANT_B.tenantId });
    (fakeInstance.getActiveAccount as ReturnType<typeof vi.fn>).mockImplementation(() => null);
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_A,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    const token = await getAccessToken();

    expect(token).toBe('token-abc');
    expect(fakeInstance.setActiveAccount).not.toHaveBeenCalled();
  });

  it('never filters by tenantId while a tenant is selected', async () => {
    const fakeInstance = makeOneProfileInstance({ profileTenantId: TENANT_B.tenantId });
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_A,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    await getAccessToken();

    for (const call of fakeInstance.getAllAccounts.mock.calls) {
      expect(call[0]).not.toHaveProperty('tenantId');
    }
  });

  it('survives an A -> B -> A -> B switch sequence (regression)', async () => {
    const fakeInstance = makeOneProfileInstance();
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    const sequence = [TENANT_A, TENANT_B, TENANT_A, TENANT_B];
    for (const tenant of sequence) {
      useAuthStore.setState({ selectedTenant: tenant });
      await expect(getAccessToken()).resolves.toBe('token-abc');
    }

    expect(fakeInstance.acquireTokenSilent).toHaveBeenCalledTimes(4);
    sequence.forEach((tenant, i) => {
      const request = fakeInstance.acquireTokenSilent.mock.calls[i][0] as Record<string, unknown>;
      expect((request.account as { tenantId: string }).tenantId).toBe(tenant.tenantId);
      expect(request.authority).toBe(`https://login.microsoftonline.com/${tenant.tenantId}`);
    });
  });

  it('mints for the signed-in user, not the leftover cached one, when the cache holds two users', async () => {
    // The leftover account sorts first, so an unfiltered [0] would issue the
    // request against ANOTHER user's refresh token — same tenant, wrong
    // identity, which the response-tenant assertion cannot catch.
    const leftover = { homeAccountId: 'home.leftover-user', tenantId: TENANT_A.tenantId, username: 'x@example.test' };
    const signedIn = { homeAccountId: HOME_ACCOUNT_ID, tenantId: TENANT_A.tenantId, username: 'y@example.test' };
    const fakeInstance = makeTwoUserInstance([leftover, signedIn]);
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_B,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    await expect(getAccessToken()).resolves.toBe('token-abc');

    const request = fakeInstance.acquireTokenSilent.mock.calls[0][0] as Record<string, unknown>;
    const account = request.account as { homeAccountId: string; username: string; tenantId: string };
    expect(account.homeAccountId).toBe(HOME_ACCOUNT_ID);
    expect(account.username).toBe('y@example.test');
    expect(account.tenantId).toBe(TENANT_B.tenantId);
  });

  it('stamps the selected tenant whichever tenant profile the cache lists first', async () => {
    // Real getAllAccounts expands ONE account entity into one AccountInfo per
    // tenant profile, so [0] is an arbitrary profile of the right account.
    // The tenant must come from the selection, not from that profile.
    const profileA = { homeAccountId: HOME_ACCOUNT_ID, tenantId: TENANT_A.tenantId, username: 'y@example.test' };
    const profileB = { homeAccountId: HOME_ACCOUNT_ID, tenantId: TENANT_B.tenantId, username: 'y@example.test' };
    for (const order of [[profileA, profileB], [profileB, profileA]]) {
      const fakeInstance = makeTwoUserInstance(order);
      useAuthStore.setState({
        msalInstance: fakeInstance as unknown as PublicClientApplication,
        selectedTenant: TENANT_B,
        homeAccountId: HOME_ACCOUNT_ID,
      });

      await expect(getAccessToken()).resolves.toBe('token-abc');

      const request = fakeInstance.acquireTokenSilent.mock.calls[0][0] as Record<string, unknown>;
      expect((request.account as { tenantId: string }).tenantId).toBe(TENANT_B.tenantId);
    }
  });

  it('rejects with "No active account" when no homeAccountId is remembered', async () => {
    // Fails closed rather than guessing from the cache: another cached user
    // would answer for the right tenant and the response assertion could not
    // catch it.
    const fakeInstance = makeOneProfileInstance();
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_A,
      homeAccountId: null,
    });

    await expect(getAccessToken()).rejects.toThrow('No active account');
    expect(fakeInstance.acquireTokenSilent).not.toHaveBeenCalled();
  });

  it('rejects with "No active account" when the remembered account is gone from the cache', async () => {
    const fakeInstance = makeOneProfileInstance();
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_A,
      homeAccountId: 'home.someone-else',
    });

    await expect(getAccessToken()).rejects.toThrow('No active account');
    expect(fakeInstance.acquireTokenSilent).not.toHaveBeenCalled();
  });
});

describe('getAccessToken — upstream path (no tenant selected)', () => {
  it('sends the upstream request shape — no authority, no cache policy', async () => {
    const account = makeAccount(TENANT_A.tenantId);
    const fakeInstance = makeUpstreamInstance({ getActiveAccount: () => account });
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: null,
    });

    await getAccessToken();

    const request = fakeInstance.acquireTokenSilent.mock.calls[0][0] as Record<string, unknown>;
    expect(Object.keys(request).sort()).toEqual(['account', 'scopes']);
    expect(request.account).toBe(account);
  });

  it('falls back to getAllAccounts()[0] with no filter and sets it active', async () => {
    const account = makeAccount(TENANT_A.tenantId);
    const fakeInstance = makeUpstreamInstance({ getAllAccounts: () => [account] });
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: null,
    });

    await getAccessToken();

    expect(fakeInstance.getAllAccounts).toHaveBeenCalledWith();
    expect(fakeInstance.setActiveAccount).toHaveBeenCalledWith(account);
  });

  it('rejects with "No active account" when the cache holds nothing', async () => {
    const fakeInstance = makeUpstreamInstance();
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: null,
    });

    await expect(getAccessToken()).rejects.toThrow('No active account');
    expect(fakeInstance.acquireTokenSilent).not.toHaveBeenCalled();
  });
});

describe('getAccessToken — response tenant assertion', () => {
  it('rejects without returning the token when the response is for a different tenant', async () => {
    // With the home account's tenant overwritten in the request, this is the
    // only cross-tenant guard left.
    const fakeInstance = makeOneProfileInstance({
      acquireTokenSilent: async () => ({ accessToken: 'token-abc', tenantId: TENANT_A.tenantId }),
    });
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_B,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    await expect(getAccessToken()).rejects.toThrow(TENANT_MISMATCH_ERROR);
    expect(fakeInstance.acquireTokenRedirect).not.toHaveBeenCalled();
  });

  it('rejects when a freshly redeemed response carries no tenantId — the guard fails closed', async () => {
    const fakeInstance = makeOneProfileInstance({
      acquireTokenSilent: async () => ({ accessToken: 'token-abc', fromCache: false }),
    });
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_B,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    await expect(getAccessToken()).rejects.toThrow(TENANT_MISMATCH_ERROR);
    expect(fakeInstance.acquireTokenRedirect).not.toHaveBeenCalled();
  });

  it('accepts a cache hit whose tenantId is empty — the lookup was realm-scoped by the request', async () => {
    // MSAL drops the realm's ID token (and answers tenantId "") when it finds
    // more than one candidate, while returning that realm's access token. The
    // cache lookup keyed on request.account.tenantId, i.e. the selection, so
    // the token IS the selected realm's; calling it a cross-tenant token
    // would name a security event that did not happen.
    const fakeInstance = makeOneProfileInstance({
      acquireTokenSilent: async () => ({ accessToken: 'token-abc', tenantId: '', fromCache: true }),
    });
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_B,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    await expect(getAccessToken()).resolves.toBe('token-abc');
  });

  it('still rejects a cache hit that names a different tenant', async () => {
    const fakeInstance = makeOneProfileInstance({
      acquireTokenSilent: async () => ({ accessToken: 'token-abc', tenantId: TENANT_A.tenantId, fromCache: true }),
    });
    useAuthStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_B,
      homeAccountId: HOME_ACCOUNT_ID,
    });

    await expect(getAccessToken()).rejects.toThrow(TENANT_MISMATCH_ERROR);
  });
});

describe('getAccessToken — interaction required', () => {
  it('shows the banner and starts no login flow when the selected tenant refuses a silent token', async () => {
    const { getAccessToken: freshGetAccessToken, authStore } = await importFreshAuth();
    const fakeInstance = makeOneProfileInstance({
      acquireTokenSilent: async () => {
        throw new InteractionRequiredAuthError('interaction_required');
      },
    });
    authStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_B,
      homeAccountId: HOME_ACCOUNT_ID,
      authError: null,
    });

    await expect(freshGetAccessToken()).rejects.toThrow('Interactive sign-in required for the selected tenant');

    expect(fakeInstance.acquireTokenRedirect).not.toHaveBeenCalled();
    const authError = authStore.getState().authError;
    expect(authError).toContain(TENANT_B.name);
    expect(authError).not.toContain(TENANT_B.tenantId);
    expect(authError).not.toContain('interaction_required');
    // The old advice ("sign out, then sign in to X") pointed at a surface
    // that no longer exists — sign-in always targets the home tenant — and
    // looped straight back into this refusal.
    expect(authError).not.toContain('Sign out');
    // What the banner offers instead: an interactive sign-in at THIS tenant.
    expect(authStore.getState().reauthTenant).toEqual(TENANT_B);
  });

  it('replaces MSAL server-error text with fixed text when a tenant is selected, and starts no login flow', async () => {
    // AADSTS500011 (the app's service principal is not in the target tenant)
    // is the likely cross-tenant failure and is not interaction-required. Its
    // text quotes the tenant id and the authority URL, and the policy store
    // renders whatever message reaches it.
    const { getAccessToken: freshGetAccessToken, authStore } = await importFreshAuth();
    const serverError = new Error(
      `AADSTS500011: The resource principal was not found in the tenant named ${TENANT_B.tenantId}. ` +
        `Trace https://login.microsoftonline.com/${TENANT_B.tenantId}/oauth2/v2.0/token`,
    );
    const fakeInstance = makeOneProfileInstance({
      acquireTokenSilent: async () => {
        throw serverError;
      },
    });
    authStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: TENANT_B,
      homeAccountId: HOME_ACCOUNT_ID,
      authError: null,
    });

    const rejection = await freshGetAccessToken().catch((e: unknown) => e);
    const message = (rejection as Error).message;

    expect(message).not.toContain(TENANT_B.tenantId);
    expect(message).not.toContain('AADSTS500011');
    expect(message).not.toContain('login.microsoftonline.com');
    expect(message).toContain(TENANT_B.name);
    expect(fakeInstance.acquireTokenRedirect).not.toHaveBeenCalled();
  });

  it('rethrows the MSAL error untouched when no tenant is selected (upstream behaviour)', async () => {
    const { getAccessToken: freshGetAccessToken, authStore } = await importFreshAuth();
    const account = makeAccount(TENANT_A.tenantId);
    const serverError = new Error('upstream_server_error');
    const fakeInstance = makeUpstreamInstance({
      getActiveAccount: () => account,
      acquireTokenSilent: async () => {
        throw serverError;
      },
    });
    authStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: null,
    });

    await expect(freshGetAccessToken()).rejects.toBe(serverError);
  });

  it('still redirects for a lapsed session when no tenant is selected', async () => {
    const { getAccessToken: freshGetAccessToken, authStore } = await importFreshAuth();
    const account = makeAccount(TENANT_A.tenantId);
    const fakeInstance = makeUpstreamInstance({
      getActiveAccount: () => account,
      acquireTokenSilent: async () => {
        throw new InteractionRequiredAuthError('interaction_required');
      },
    });
    authStore.setState({
      msalInstance: fakeInstance as unknown as PublicClientApplication,
      selectedTenant: null,
    });

    await expect(freshGetAccessToken()).rejects.toThrow('Redirecting to login');

    expect(fakeInstance.acquireTokenRedirect).toHaveBeenCalledTimes(1);
    expect(authStore.getState().authError).toBeNull();
  });
});
