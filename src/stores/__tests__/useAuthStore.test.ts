// stores/__tests__/useAuthStore.test.ts — Authority construction and
// persist-before-redirect ordering.
//
// @azure/msal-browser is mocked because the real module needs a browser.
// The ordering assertion (persist before loginRedirect) exists because
// loginRedirect unloads the page — anything written after it is too late.

import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

type LoggerCallback = (level: number, message: string, containsPii: boolean) => void;
const constructedConfigs: Array<{
  auth: { authority: string };
  system?: { loggerOptions: { loggerCallback: LoggerCallback } };
}> = [];
const callOrder: string[] = [];

interface MockAccount {
  homeAccountId: string;
  tenantId: string;
  username: string;
  /** Mirrors MSAL's isHomeTenant filter (tid === utid). Defaults to true. */
  isHomeTenant?: boolean;
}

// Driven per-test to simulate accounts already present in MSAL's cache
// before initializeMsal() resolves them.
let mockAccounts: MockAccount[] = [];
let mockActiveAccount: MockAccount | null = null;
// When set, the next instance's initialize() rejects with this error.
let mockInitializeError: Error | null = null;
// When set, handleRedirectPromise() fires LOGIN_SUCCESS for this account
// through the registered event callbacks — the real ordering: the callback
// runs INSIDE handleRedirectPromise, before initializeMsal's final set().
let mockLoginSuccessAccount: MockAccount | null = null;

vi.mock('@azure/msal-browser', () => {
  class PublicClientApplication {
    config: { auth: { authority: string } };
    constructor(config: { auth: { authority: string } }) {
      this.config = config;
      constructedConfigs.push(config);
    }
    initialize = vi.fn(async () => {
      if (mockInitializeError) throw mockInitializeError;
    });
    callbacks: Array<(event: unknown) => void> = [];
    addEventCallback = vi.fn((cb: (event: unknown) => void) => {
      this.callbacks.push(cb);
      return 'callback-id';
    });
    handleRedirectPromise = vi.fn(async () => {
      if (mockLoginSuccessAccount) {
        const account = mockLoginSuccessAccount;
        for (const cb of this.callbacks) {
          cb({ eventType: 'msal:loginSuccess', payload: { account } });
        }
      }
      return null;
    });
    getActiveAccount = vi.fn(() => mockActiveAccount);
    getAllAccounts = vi.fn(
      (filter?: { tenantId?: string; homeAccountId?: string; isHomeTenant?: boolean }) => {
        let list = mockAccounts;
        if (filter?.tenantId) list = list.filter((a) => a.tenantId === filter.tenantId);
        if (filter?.homeAccountId) list = list.filter((a) => a.homeAccountId === filter.homeAccountId);
        if (filter?.isHomeTenant) list = list.filter((a) => a.isHomeTenant !== false);
        return list;
      },
    );
    // Stateful, unlike a bare no-op: the real MSAL setActiveAccount(null)
    // removes the active-account key outright, and setActiveAccount(acct)
    // makes getActiveAccount() return it. A no-op mock would let the
    // dead-code bug pass silently (the store's re-derivation logic depends
    // on the clear actually taking effect).
    setActiveAccount = vi.fn((a: MockAccount | null) => {
      mockActiveAccount = a;
    });
    loginRedirect = vi.fn(async () => {
      callOrder.push('loginRedirect');
    });
    acquireTokenRedirect = vi.fn(async () => {
      callOrder.push('acquireTokenRedirect');
    });
    logoutRedirect = vi.fn(async () => {});
  }

  return {
    PublicClientApplication,
    EventType: { LOGIN_SUCCESS: 'msal:loginSuccess' },
    LogLevel: { Error: 0, Warning: 1, Info: 2, Verbose: 3, Trace: 4 },
    BrowserAuthError: class BrowserAuthError extends Error {},
    BrowserAuthErrorCodes: {},
    InteractionRequiredAuthError: class InteractionRequiredAuthError extends Error {},
  };
});

import { useAuthStore } from '../useAuthStore';
import { DEFAULT_AUTHORITY, TENANT_CHOICE_KEY, resolveTenant, type TenantEntry } from '../../lib/tenants';
import { usePolicyStore } from '../usePolicyStore';
import { usePersonaStore } from '../usePersonaStore';
import { useEvaluationStore } from '../useEvaluationStore';

const TENANT_A: TenantEntry = { name: 'Tenant A', tenantId: '11111111-1111-1111-1111-111111111111' };
const TENANT_B: TenantEntry = { name: 'Tenant B', tenantId: '22222222-2222-2222-2222-222222222222' };
const TENANT_C: TenantEntry = { name: 'Tenant C', tenantId: '33333333-3333-3333-3333-333333333333' };

function stubStorage(): Map<string, string> {
  const store = new Map<string, string>();
  vi.stubGlobal('localStorage', {
    getItem: (k: string) => store.get(k) ?? null,
    setItem: (k: string, v: string) => {
      store.set(k, v);
      callOrder.push('setItem');
    },
    removeItem: (k: string) => void store.delete(k),
  });
  return store;
}

// switchInProgress is a module-scope latch in useAuthStore.ts guarding the
// in-flight window, and bootPromise is a second module-scope value. Cases that
// depend on either take a fresh module instance so no state leaks between them
// — vi.resetModules() + a dynamic re-import per case, including the three data
// stores useAuthStore clears, since a fresh useAuthStore module closes over
// fresh copies of those imports too.
async function importFreshAuthStore() {
  vi.resetModules();
  const [authMod, policyMod, personaMod, evalMod] = await Promise.all([
    import('../useAuthStore'),
    import('../usePolicyStore'),
    import('../usePersonaStore'),
    import('../useEvaluationStore'),
  ]);
  return {
    auth: authMod.useAuthStore,
    policy: policyMod.usePolicyStore,
    persona: personaMod.usePersonaStore,
    evaluation: evalMod.useEvaluationStore,
  };
}

/**
 * Fails fast instead of hanging: every promise chain asserted here is expected
 * to settle on its own, so a chain that cannot (a boot promise awaiting a
 * switch that awaits the boot promise) reports the deadlock rather than the
 * suite's generic timeout.
 */
function settlesWithin<T>(promise: Promise<T>, ms = 200): Promise<T> {
  return Promise.race([
    promise,
    new Promise<never>((_, reject) =>
      setTimeout(() => reject(new Error('promise did not settle — deadlock')), ms),
    ),
  ]);
}

beforeEach(() => {
  constructedConfigs.length = 0;
  callOrder.length = 0;
  mockAccounts = [];
  mockActiveAccount = null;
  mockInitializeError = null;
  mockLoginSuccessAccount = null;
  vi.stubEnv('VITE_MSAL_CLIENT_ID', 'test-client-id');
  vi.stubGlobal('window', { location: { origin: 'https://example.test' } });
  stubStorage();
  useAuthStore.setState({
    msalInstance: null,
    isInitialized: false,
    isInitializing: false,
    authError: null,
    accounts: [],
    isAuthenticated: false,
    selectedTenant: null,
    reauthTenant: null,
    isSwitching: false,
    homeAccountId: null,
    tenants: [],
  });
  useAuthStore.getState().setBootPromise(null);
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.unstubAllEnvs();
  vi.clearAllMocks();
});

describe('useAuthStore.login', () => {
  // Sign-in always targets the account's home tenant via the default
  // authority: login() takes no tenant, chooses
  // none and persists none — a remembered or deep-linked tenant is applied
  // afterwards as an automatic switch.
  it('constructs with DEFAULT_AUTHORITY, writes nothing to storage, and redirects', async () => {
    await useAuthStore.getState().login();
    expect(constructedConfigs).toHaveLength(1);
    expect(constructedConfigs[0].auth.authority).toBe(DEFAULT_AUTHORITY);
    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBeNull();
    expect(callOrder).toEqual(['loginRedirect']);
  });

  it('still uses the default authority when a tenant is already selected from a previous switch', async () => {
    useAuthStore.setState({ selectedTenant: TENANT_B });

    await useAuthStore.getState().login();

    expect(constructedConfigs).toHaveLength(1);
    expect(constructedConfigs[0].auth.authority).toBe(DEFAULT_AUTHORITY);
    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBeNull();
    expect(callOrder).toEqual(['loginRedirect']);
  });

  it('adopts a deep-linked ?tenant= only at sign-in, and before the redirect unloads the page', async () => {
    // Persisting the parameter at boot would let a shared link rewrite the
    // remembered tenant of anyone who merely opened it. Signing in is the act
    // that adopts it, and the return leg reads storage only — so the write
    // has to land before loginRedirect.
    vi.stubGlobal('window', {
      location: { origin: 'https://example.test', search: `?tenant=${TENANT_B.tenantId}` },
    });
    useAuthStore.getState().setTenants([TENANT_A, TENANT_B]);

    await useAuthStore.getState().login();

    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBe(TENANT_B.tenantId);
    expect(callOrder).toEqual(['setItem', 'loginRedirect']);
  });

  it('writes nothing for a ?tenant= value that is not in the curated list', async () => {
    vi.stubGlobal('window', {
      location: { origin: 'https://example.test', search: `?tenant=${TENANT_C.tenantId}` },
    });
    useAuthStore.getState().setTenants([TENANT_A, TENANT_B]);

    await useAuthStore.getState().login();

    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBeNull();
    expect(callOrder).toEqual(['loginRedirect']);
  });

  it('does not redirect when a session is already live', async () => {
    mockAccounts = [{ homeAccountId: 'h1', tenantId: TENANT_A.tenantId, username: 'a@a' }];
    await useAuthStore.getState().initializeMsal(null);
    expect(useAuthStore.getState().isAuthenticated).toBe(true);
    const instance = useAuthStore.getState().msalInstance!;
    const loginRedirectSpy = instance.loginRedirect as ReturnType<typeof vi.fn>;
    loginRedirectSpy.mockClear();

    await useAuthStore.getState().login();

    expect(loginRedirectSpy).not.toHaveBeenCalled();
    expect(constructedConfigs).toHaveLength(1);
  });

  it('rebuilds a signed-out instance that was built for a tenant, at the default authority', async () => {
    // The boot effect initialised for a tenant (leftover msal.* keys after a
    // logout), then the user clicks the plain Sign in button.
    await useAuthStore.getState().initializeMsal(TENANT_A);
    expect(constructedConfigs[0].auth.authority).toBe(`https://login.microsoftonline.com/${TENANT_A.tenantId}`);
    expect(useAuthStore.getState().isAuthenticated).toBe(false);

    await useAuthStore.getState().login();

    expect(constructedConfigs).toHaveLength(2);
    expect(constructedConfigs[1].auth.authority).toBe(DEFAULT_AUTHORITY);
    expect(useAuthStore.getState().authError).toBeNull();
    expect(useAuthStore.getState().selectedTenant).toBeNull();
    expect(callOrder).toEqual(['loginRedirect']);
  });

  it('awaits an in-flight boot initialization and redirects on that instance', async () => {
    // Boot's initializeMsal(null) is started but not yet awaited — mirrors
    // App.tsx's boot effect racing a user click on "Sign In".
    const bootInit = useAuthStore.getState().initializeMsal(null);

    await useAuthStore.getState().login();
    await bootInit;

    // Both legs target the default authority now, so the boot instance is the
    // one that redirects; no second instance is built.
    expect(constructedConfigs).toHaveLength(1);
    expect(constructedConfigs[0].auth.authority).toBe(DEFAULT_AUTHORITY);
    expect(callOrder).toEqual(['loginRedirect']);
  });

  it('waits for a boot leg that has not started init yet, then does not redirect over a live session', async () => {
    // App.tsx registers the boot promise before tenants.json resolves. A
    // click on the plain Sign in button lands in that window: tenants is
    // still [], no instance exists, and the redirect params are pending.
    mockAccounts = [{ homeAccountId: 'h1', tenantId: TENANT_A.tenantId, username: 'a@a' }];
    let finishFetch!: () => void;
    const fetchDone = new Promise<void>((resolve) => { finishFetch = resolve; });
    useAuthStore.getState().setBootPromise(
      fetchDone.then(() => useAuthStore.getState().initializeMsal(null)),
    );

    const click = useAuthStore.getState().login();
    // Without the settle step, login() would already have built a second
    // instance here and consumed the pending redirect.
    await Promise.resolve();
    expect(constructedConfigs).toHaveLength(0);

    finishFetch();
    await click;

    expect(constructedConfigs).toHaveLength(1);
    expect(useAuthStore.getState().isAuthenticated).toBe(true);
    // The boot leg completed the redirect return, so the stale click must
    // not start a second redirect.
    expect(callOrder).not.toContain('loginRedirect');
  });
});

describe('useAuthStore.homeAccountId', () => {
  it('initializes to null', () => {
    expect(useAuthStore.getState().homeAccountId).toBeNull();
  });

  it('is set from the unfiltered account list at boot', async () => {
    mockAccounts = [{ homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'a@a' }];

    await useAuthStore.getState().initializeMsal(TENANT_B);

    expect(useAuthStore.getState().homeAccountId).toBe('home.signed-in');
  });

  it('names the account that just signed in, not a leftover cached user', async () => {
    // The tab still holds another user's account (MSAL keeps msal.* keys
    // after an abandoned logout) and it sorts first. LOGIN_SUCCESS fires
    // inside handleRedirectPromise for the user who actually signed in, so
    // the final init write must not replace that id with getAllAccounts()[0]
    // — a token minted for the wrong user passes the response-tenant
    // assertion untouched (right tenant, other identity).
    const stale = { homeAccountId: 'home.leftover-user', tenantId: TENANT_A.tenantId, username: 'x@x' };
    const signedIn = { homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'y@y' };
    mockAccounts = [stale, signedIn];
    mockLoginSuccessAccount = signedIn;

    await useAuthStore.getState().initializeMsal();

    expect(useAuthStore.getState().homeAccountId).toBe('home.signed-in');
    expect(useAuthStore.getState().msalInstance!.getActiveAccount()).toEqual(signedIn);
  });

  it('falls back to the first cached account when nothing is active (page refresh)', async () => {
    mockAccounts = [{ homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'a@a' }];
    mockActiveAccount = null;

    await useAuthStore.getState().initializeMsal();

    expect(useAuthStore.getState().homeAccountId).toBe('home.signed-in');
  });

  it('is cleared by logout() and left untouched by switchTenant()', async () => {
    mockAccounts = [{ homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'a@a' }];
    await useAuthStore.getState().initializeMsal(null);
    expect(useAuthStore.getState().homeAccountId).toBe('home.signed-in');

    await useAuthStore.getState().switchTenant(TENANT_B);
    expect(useAuthStore.getState().homeAccountId).toBe('home.signed-in');

    await useAuthStore.getState().logout();
    expect(useAuthStore.getState().homeAccountId).toBeNull();
  });
});

describe('useAuthStore.initializeMsal', () => {
  it('surfaces a fixed message, never MSAL text, when initialize() rejects', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});
    // MSAL's endpoint-resolution errors quote the authority URL, which now
    // embeds the tenant id.
    mockInitializeError = new Error(
      `endpoints_resolution_error: https://login.microsoftonline.com/${TENANT_A.tenantId}`,
    );

    await useAuthStore.getState().initializeMsal(TENANT_A);

    const { authError, msalInstance, isInitialized, isInitializing } = useAuthStore.getState();
    expect(authError).toBe('Could not initialize sign-in. Check your network and try again.');
    expect(authError).not.toContain(TENANT_A.tenantId);
    expect(msalInstance).toBeNull();
    expect(isInitialized).toBe(false);
    expect(isInitializing).toBe(false);
    expect(consoleError).toHaveBeenCalled();
    for (const call of consoleError.mock.calls) {
      expect(String(call[0])).not.toContain(TENANT_A.tenantId);
    }
    consoleError.mockRestore();
  });

  it('redacts the tenant id from the lines MSAL\'s own logger relays to the console', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});
    const consoleWarn = vi.spyOn(console, 'warn').mockImplementation(() => {});

    await useAuthStore.getState().initializeMsal(TENANT_A);
    const loggerCallback = constructedConfigs[0].system!.loggerOptions.loggerCallback;

    loggerCallback(0, `endpoints_resolution_error: https://login.microsoftonline.com/${TENANT_A.tenantId}/`, false);
    loggerCallback(1, `slow authority ${TENANT_A.tenantId}`, false);
    loggerCallback(0, `pii line ${TENANT_A.tenantId}`, true);

    expect(consoleError).toHaveBeenCalledTimes(1);
    expect(consoleError.mock.calls[0][0]).toContain('endpoints_resolution_error');
    expect(consoleError.mock.calls[0][0]).not.toContain(TENANT_A.tenantId);
    expect(consoleWarn).toHaveBeenCalledTimes(1);
    expect(consoleWarn.mock.calls[0][0]).not.toContain(TENANT_A.tenantId);
    consoleError.mockRestore();
    consoleWarn.mockRestore();
  });

  it('redacts a tenant id the instance was NOT built for — the instance serves every switched-to tenant', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});

    await useAuthStore.getState().initializeMsal(TENANT_A);
    const loggerCallback = constructedConfigs[0].system!.loggerOptions.loggerCallback;

    // After an in-place switch the token request carries tenant B's authority
    // while this instance is still the one built for tenant A.
    loggerCallback(0, `invalid_grant for https://login.microsoftonline.com/${TENANT_B.tenantId}/oauth2/v2.0/token`, false);

    expect(consoleError).toHaveBeenCalledTimes(1);
    const line = String(consoleError.mock.calls[0][0]);
    expect(line).toContain('invalid_grant');
    expect(line).not.toContain(TENANT_B.tenantId);
    expect(line).toContain('<id>');
    consoleError.mockRestore();
  });

  it('relays MSAL logger lines unchanged for the default authority', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});

    await useAuthStore.getState().initializeMsal(null);
    constructedConfigs[0].system!.loggerOptions.loggerCallback(0, `error at ${DEFAULT_AUTHORITY}`, false);

    expect(consoleError).toHaveBeenCalledWith(`error at ${DEFAULT_AUTHORITY}`);
    consoleError.mockRestore();
  });

  it('sets selectedTenant to the tenant passed in', async () => {
    await useAuthStore.getState().initializeMsal(TENANT_A);
    expect(useAuthStore.getState().selectedTenant).toEqual(TENANT_A);
  });

  it('records homeAccountId and the unfiltered account list in the LOGIN_SUCCESS callback', async () => {
    const accountA = { homeAccountId: 'a', tenantId: TENANT_A.tenantId, username: 'a@a.test' };
    const accountB = { homeAccountId: 'b', tenantId: TENANT_B.tenantId, username: 'b@b.test' };
    mockAccounts = [accountB, accountA];

    await useAuthStore.getState().initializeMsal(TENANT_A);
    const instance = useAuthStore.getState().msalInstance!;
    const addEventCallbackSpy = instance.addEventCallback as ReturnType<typeof vi.fn>;
    const callback = addEventCallbackSpy.mock.calls[0][0] as (event: unknown) => void;

    callback({ eventType: 'msal:loginSuccess', payload: { account: accountA } });

    expect(useAuthStore.getState().isAuthenticated).toBe(true);
    expect(useAuthStore.getState().homeAccountId).toBe('a');
    expect(useAuthStore.getState().accounts).toEqual([accountB, accountA]);
  });

  it('leaves selectedTenant null when called with no tenant', async () => {
    await useAuthStore.getState().initializeMsal();
    expect(useAuthStore.getState().selectedTenant).toBeNull();
  });

  it('does not drop an existing selection when called with no tenant', async () => {
    // The parameter is a starting selection, never a reset: a later init that
    // cleared it would revert every token request to the home tenant while
    // the header still named the selected one.
    useAuthStore.setState({ selectedTenant: TENANT_B });

    await useAuthStore.getState().initializeMsal();

    expect(useAuthStore.getState().selectedTenant).toEqual(TENANT_B);
  });
});

describe('useAuthStore.switchTenant', () => {
  it('persists next.tenantId to localStorage', async () => {
    const { auth } = await importFreshAuthStore();
    await auth.getState().switchTenant(TENANT_A);
    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBe(TENANT_A.tenantId);
  });

  it('sets selectedTenant to the new tenant argument', async () => {
    const { auth } = await importFreshAuthStore();
    await auth.getState().switchTenant(TENANT_A);
    expect(auth.getState().selectedTenant).toEqual(TENANT_A);
  });

  it('clears a stale authError left over from the tenant being switched away from', async () => {
    const { auth } = await importFreshAuthStore();
    auth.setState({ authError: 'Sign out before signing in to a different tenant.' });
    await auth.getState().switchTenant(TENANT_A);
    expect(auth.getState().authError).toBeNull();
  });

  it('clears all three data stores and calls neither logoutRedirect nor loginRedirect', async () => {
    const { auth, policy, persona, evaluation } = await importFreshAuthStore();
    // Sign in first so an msalInstance exists — the switch must leave that
    // session alone rather than ending it.
    await auth.getState().login();

    const policyClear = vi.fn(() => callOrder.push('policyClear'));
    const personaClear = vi.fn(() => callOrder.push('personaClear'));
    const evaluationClear = vi.fn(() => callOrder.push('evaluationClear'));
    policy.setState({ clear: policyClear });
    persona.setState({ clear: personaClear });
    evaluation.setState({ clear: evaluationClear });

    const instance = auth.getState().msalInstance!;
    const loginRedirectSpy = instance.loginRedirect as ReturnType<typeof vi.fn>;
    const logoutRedirectSpy = instance.logoutRedirect as ReturnType<typeof vi.fn>;
    loginRedirectSpy.mockClear();
    callOrder.length = 0;

    await auth.getState().switchTenant(TENANT_B);

    expect(policyClear).toHaveBeenCalledTimes(1);
    expect(personaClear).toHaveBeenCalledTimes(1);
    expect(evaluationClear).toHaveBeenCalledTimes(1);
    expect(logoutRedirectSpy).not.toHaveBeenCalled();
    expect(loginRedirectSpy).not.toHaveBeenCalled();
    // setItem (persistTenant) fires before any store is cleared.
    expect(callOrder).toEqual(['setItem', 'policyClear', 'personaClear', 'evaluationClear']);
  });

  it('leaves the signed-in session visible: isAuthenticated and accounts are untouched', async () => {
    const { auth } = await importFreshAuthStore();
    mockAccounts = [{ homeAccountId: 'h1', tenantId: TENANT_A.tenantId, username: 'a@a' }];
    await auth.getState().login();
    const accountsBefore = auth.getState().accounts;
    expect(auth.getState().isAuthenticated).toBe(true);

    await auth.getState().switchTenant(TENANT_B);

    // No account exists for the new tenant until its first token, so
    // re-deriving here would sign the user out visually for no reason.
    expect(auth.getState().isAuthenticated).toBe(true);
    expect(auth.getState().accounts).toBe(accountsBefore);
  });

  it('reloads the new tenant in place when live data was loaded before the switch', async () => {
    const { auth, policy } = await importFreshAuthStore();
    await auth.getState().login();
    const loadFromGraph = vi.fn(async () => {});
    policy.setState({ dataSource: 'live', loadFromGraph, clear: vi.fn() });

    await auth.getState().switchTenant(TENANT_B);

    expect(loadFromGraph).toHaveBeenCalledTimes(1);
  });

  it('does not reload when the previous data source was sample or none', async () => {
    const { auth, policy } = await importFreshAuthStore();
    await auth.getState().login();
    const loadFromGraph = vi.fn(async () => {});
    policy.setState({ dataSource: 'sample', loadFromGraph, clear: vi.fn() });

    await auth.getState().switchTenant(TENANT_B);
    expect(loadFromGraph).not.toHaveBeenCalled();

    policy.setState({ dataSource: 'none' });
    await auth.getState().switchTenant(TENANT_A);
    expect(loadFromGraph).not.toHaveBeenCalled();
  });

  it('keeps the new tenant selected when the reload fails, and does not throw', async () => {
    const { auth, policy } = await importFreshAuthStore();
    await auth.getState().login();
    const loadFromGraph = vi.fn(async () => {
      throw new Error('graph failed: https://login.microsoftonline.com/secret-tenant');
    });
    policy.setState({ dataSource: 'live', loadFromGraph, clear: vi.fn() });

    await expect(auth.getState().switchTenant(TENANT_B)).resolves.toBeUndefined();

    // The switch itself succeeded; the policy store owns the load error.
    expect(auth.getState().selectedTenant).toEqual(TENANT_B);
    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBe(TENANT_B.tenantId);
  });

  it('latches only the critical section: a concurrent call no-ops, a later one switches again', async () => {
    const { auth } = await importFreshAuthStore();
    await auth.getState().login();
    let finishBoot!: () => void;
    const gate = new Promise<void>((resolve) => { finishBoot = resolve; });
    auth.getState().setBootPromise(gate);

    const first = auth.getState().switchTenant(TENANT_B);
    // Second click while the first switch is still settling. It targets a
    // THIRD tenant so "the second call no-oped" is distinguishable from "the
    // second call wrote the tenant we were already on".
    await auth.getState().switchTenant(TENANT_C);
    // Sign-in selects no tenant, so a dropped second switch leaves the
    // selection exactly where sign-in left it: null.
    expect(auth.getState().selectedTenant).toBeNull();

    finishBoot();
    await first;
    expect(auth.getState().selectedTenant).toEqual(TENANT_B);

    // The page no longer unloads, so the latch must be released for the
    // next switch rather than making it a silent no-op.
    await auth.getState().switchTenant(TENANT_C);
    expect(auth.getState().selectedTenant).toEqual(TENANT_C);
  });

  it('reports isSwitching for the critical section only', async () => {
    const { auth } = await importFreshAuthStore();
    await auth.getState().login();
    let finishBoot!: () => void;
    const gate = new Promise<void>((resolve) => { finishBoot = resolve; });
    auth.getState().setBootPromise(gate);

    expect(auth.getState().isSwitching).toBe(false);
    const first = auth.getState().switchTenant(TENANT_B);
    // The window in which a second pick would be dropped by the latch: the
    // Switch tenant items are disabled for exactly this long.
    expect(auth.getState().isSwitching).toBe(true);

    finishBoot();
    await first;
    expect(auth.getState().isSwitching).toBe(false);
  });

  it('clears isSwitching even when the switch takes the no-instance branch', async () => {
    const { auth } = await importFreshAuthStore();
    expect(auth.getState().msalInstance).toBeNull();

    await auth.getState().switchTenant(TENANT_A);

    expect(auth.getState().isSwitching).toBe(false);
  });

  it('accepts a switch requested while the previous switch is still reloading', async () => {
    const { auth, policy } = await importFreshAuthStore();
    await auth.getState().login();
    const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

    let finishLoad!: () => void;
    const slowLoad = new Promise<void>((resolve) => { finishLoad = resolve; });
    const loadFromGraph = vi.fn(async () => {
      await slowLoad;
    });
    policy.setState({ dataSource: 'live', loadFromGraph, clear: vi.fn() });

    const first = auth.getState().switchTenant(TENANT_B);
    await flush();
    expect(loadFromGraph).toHaveBeenCalledTimes(1);
    expect(auth.getState().selectedTenant).toEqual(TENANT_B);

    // The latch must already be released: the reload is not part of the
    // critical section, so this second switch has to be accepted rather than
    // dropped with no state change and no feedback.
    const second = auth.getState().switchTenant(TENANT_C);
    await flush();
    expect(auth.getState().selectedTenant).toEqual(TENANT_C);
    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBe(TENANT_C.tenantId);

    finishLoad();
    await Promise.all([first, second]);
    expect(auth.getState().selectedTenant).toEqual(TENANT_C);
  });

  it('with msalInstance null persists the choice, clears the stores, and calls no redirect', async () => {
    const { auth, policy, persona, evaluation } = await importFreshAuthStore();

    const policyClear = vi.fn();
    const personaClear = vi.fn();
    const evaluationClear = vi.fn();
    policy.setState({ clear: policyClear });
    persona.setState({ clear: personaClear });
    evaluation.setState({ clear: evaluationClear });

    expect(auth.getState().msalInstance).toBeNull();
    await auth.getState().switchTenant(TENANT_A);

    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBe(TENANT_A.tenantId);
    expect(policyClear).toHaveBeenCalledTimes(1);
    expect(personaClear).toHaveBeenCalledTimes(1);
    expect(evaluationClear).toHaveBeenCalledTimes(1);
    expect(constructedConfigs).toHaveLength(0);
  });

  it('awaits a boot leg that has not started init yet before reading msalInstance', async () => {
    const { auth } = await importFreshAuthStore();
    mockAccounts = [{ homeAccountId: 'h1', tenantId: TENANT_A.tenantId, username: 'a@a' }];
    let finishFetch!: () => void;
    const fetchDone = new Promise<void>((resolve) => { finishFetch = resolve; });
    auth.getState().setBootPromise(fetchDone.then(() => auth.getState().initializeMsal(TENANT_A)));

    const switching = auth.getState().switchTenant(TENANT_B);
    await Promise.resolve();
    expect(auth.getState().msalInstance).toBeNull();
    finishFetch();
    await switching;

    const instance = auth.getState().msalInstance;
    expect(instance).not.toBeNull();
    expect(constructedConfigs[0].auth.authority).toBe(`https://login.microsoftonline.com/${TENANT_A.tenantId}`);
    expect(instance!.logoutRedirect).not.toHaveBeenCalled();
    expect(auth.getState().selectedTenant).toEqual(TENANT_B);
  });

  it('awaits an in-flight boot initialization before reading msalInstance', async () => {
    const { auth } = await importFreshAuthStore();
    // Boot init is started but not yet awaited when the switch arrives.
    const bootInit = auth.getState().initializeMsal(TENANT_A);

    await auth.getState().switchTenant(TENANT_B);
    await bootInit;

    // Without the settle step, switchTenant would have taken the no-instance
    // branch (msalInstance still null at that point), so the live data reload
    // decision would be made before boot had produced an instance at all.
    const instance = auth.getState().msalInstance;
    expect(instance).not.toBeNull();
    expect(instance!.logoutRedirect).not.toHaveBeenCalled();
    expect(auth.getState().selectedTenant).toEqual(TENANT_B);
  });
});

// The boot leg lives in App.tsx, which this project does not component-test.
// Its promise TOPOLOGY is plain promises over this store, though, and that is
// what deadlocked: registering the whole chain — including the automatic first
// switch — as the boot promise made switchTenant's `await settleBoot()` wait
// for a promise that could not settle until the switch returned. These cases
// reproduce the exact shape of App.tsx's redirect-return leg.
describe('App boot leg topology — automatic first switch', () => {
  async function runBootLeg(auth: typeof useAuthStore, tenants: TenantEntry[], persistedId: string | null) {
    const boot = Promise.resolve(tenants).then(async (list) => {
      auth.getState().setTenants(list);
      const resolved = resolveTenant(list, { urlParam: null, persistedId });
      // isAuthenticated flips inside initializeMsal, revealing the Switch
      // tenant items before the automatic switch has run.
      if (resolved) auth.setState({ isSwitching: true });
      await auth.getState().initializeMsal();
      return resolved;
    });
    // Only up to and including initializeMsal — never the switch below.
    auth.getState().setBootPromise(boot.then(() => undefined));
    return boot.then((resolved) => {
      const { isAuthenticated, selectedTenant, switchTenant } = auth.getState();
      if (resolved && isAuthenticated && selectedTenant === null) {
        return switchTenant(resolved);
      }
      auth.setState({ isSwitching: false });
    });
  }

  it('settles, applies the remembered tenant, and leaves the latch open for a later switch', async () => {
    const { auth } = await importFreshAuthStore();
    mockAccounts = [{ homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'a@a' }];

    await settlesWithin(runBootLeg(auth, [TENANT_A, TENANT_B], TENANT_B.tenantId));

    expect(auth.getState().selectedTenant).toEqual(TENANT_B);
    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBe(TENANT_B.tenantId);
    // The submenu is usable again once the automatic switch has landed.
    expect(auth.getState().isSwitching).toBe(false);
    // The deadlock also stuck switchInProgress at true for the life of the
    // page, so every later user switch was dropped in silence.
    await settlesWithin(auth.getState().switchTenant(TENANT_C));
    expect(auth.getState().selectedTenant).toEqual(TENANT_C);
  });

  it('leaves login() usable — it awaits the same boot promise', async () => {
    const { auth } = await importFreshAuthStore();
    mockAccounts = [{ homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'a@a' }];

    await settlesWithin(runBootLeg(auth, [TENANT_A, TENANT_B], TENANT_B.tenantId));
    await settlesWithin(auth.getState().login());

    // A live session, so login() returns without redirecting — the point is
    // that it returns at all.
    expect(callOrder).not.toContain('loginRedirect');
  });

  it('switches exactly once when the effect runs twice (React StrictMode)', async () => {
    const { auth, policy } = await importFreshAuthStore();
    mockAccounts = [{ homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'a@a' }];
    const clear = vi.fn();
    policy.setState({ clear });

    await settlesWithin(
      Promise.all([
        runBootLeg(auth, [TENANT_A, TENANT_B], TENANT_B.tenantId),
        runBootLeg(auth, [TENANT_A, TENANT_B], TENANT_B.tenantId),
      ]),
    );

    expect(auth.getState().selectedTenant).toEqual(TENANT_B);
    expect(clear).toHaveBeenCalledTimes(1);
  });

  it('does nothing when no tenant resolves', async () => {
    const { auth } = await importFreshAuthStore();
    mockAccounts = [{ homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'a@a' }];

    await settlesWithin(runBootLeg(auth, [TENANT_A, TENANT_B], null));

    expect(auth.getState().selectedTenant).toBeNull();
    expect(localStorage.getItem(TENANT_CHOICE_KEY)).toBeNull();
    expect(auth.getState().isSwitching).toBe(false);
  });
});

// The cure for a tenant's token refusal. The refusal is usually the target
// tenant's sign-in frequency policy rejecting the silent redeem on auth_time
// (AADSTS70043), so the recovery has to be an interactive sign-in AT that tenant
// — and one the user starts.
describe('useAuthStore.reauthenticate', () => {
  it('redirects interactively to the tenant authority with prompt=login, and starts no login flow', async () => {
    const { auth } = await importFreshAuthStore();
    mockAccounts = [{ homeAccountId: 'home.signed-in', tenantId: TENANT_A.tenantId, username: 'a@a' }];
    await auth.getState().initializeMsal();
    auth.setState({ authError: 'Could not get access to Tenant B without signing in again.', reauthTenant: TENANT_B });
    callOrder.length = 0;

    await auth.getState().reauthenticate(TENANT_B);

    const instance = auth.getState().msalInstance!;
    const redirectSpy = (instance as unknown as { acquireTokenRedirect: ReturnType<typeof vi.fn> }).acquireTokenRedirect;
    expect(redirectSpy).toHaveBeenCalledTimes(1);
    const request = redirectSpy.mock.calls[0][0] as Record<string, unknown>;
    expect(request.authority).toBe(`https://login.microsoftonline.com/${TENANT_B.tenantId}`);
    expect(request.prompt).toBe('login');
    expect(request.scopes).toBeDefined();
    expect(instance.loginRedirect).not.toHaveBeenCalled();
    expect(instance.logoutRedirect).not.toHaveBeenCalled();
    // The banner goes away with the act that supersedes it, and the
    // remembered tenant is left alone so the automatic first switch lands
    // back in it after the return.
    expect(auth.getState().authError).toBeNull();
    expect(auth.getState().reauthTenant).toBeNull();
    expect(callOrder).toEqual(['acquireTokenRedirect']);
  });

  it('does nothing without an instance', async () => {
    const { auth } = await importFreshAuthStore();
    await expect(auth.getState().reauthenticate(TENANT_B)).resolves.toBeUndefined();
    expect(constructedConfigs).toHaveLength(0);
  });

  it('is cleared with the banner it belongs to', async () => {
    useAuthStore.setState({ authError: 'x', reauthTenant: TENANT_B });
    useAuthStore.getState().clearAuthError();
    expect(useAuthStore.getState().reauthTenant).toBeNull();
  });

  it('is cleared by a switch to another tenant', async () => {
    const { auth } = await importFreshAuthStore();
    auth.setState({ authError: 'x', reauthTenant: TENANT_B });
    await auth.getState().switchTenant(TENANT_C);
    expect(auth.getState().reauthTenant).toBeNull();
  });
});

describe('tenants / setTenants', () => {
  it('initializes to an empty array', () => {
    expect(useAuthStore.getState().tenants).toEqual([]);
  });

  it('replaces the tenant list with the entries passed in', () => {
    useAuthStore.getState().setTenants([TENANT_A, TENANT_B]);
    expect(useAuthStore.getState().tenants).toEqual([TENANT_A, TENANT_B]);
  });

  it('leaves tenants populated after logout() — the list is configuration, not session state', async () => {
    useAuthStore.getState().setTenants([TENANT_A, TENANT_B]);
    await useAuthStore.getState().login();
    await useAuthStore.getState().logout();
    expect(useAuthStore.getState().tenants).toEqual([TENANT_A, TENANT_B]);
  });
});

describe('useAuthStore.initializeMsal — account resolution', () => {
  it('calls getAllAccounts with no filter argument when initialized with null', async () => {
    await useAuthStore.getState().initializeMsal(null);
    const instance = useAuthStore.getState().msalInstance;
    const getAllAccountsSpy = instance!.getAllAccounts as ReturnType<typeof vi.fn>;
    expect(getAllAccountsSpy).toHaveBeenCalledWith();
  });

  it('calls getAllAccounts with no filter argument even when a tenant is selected', async () => {
    // MSAL keeps only the most recent tenant profile, so a tenantId filter here
    // answers empty for the tenant switched away from and would sign the user
    // out visually.
    mockAccounts = [{ homeAccountId: 'home.a', tenantId: TENANT_A.tenantId, username: 'a@a' }];

    await useAuthStore.getState().initializeMsal(TENANT_B);

    const instance = useAuthStore.getState().msalInstance;
    const getAllAccountsSpy = instance!.getAllAccounts as ReturnType<typeof vi.fn>;
    for (const call of getAllAccountsSpy.mock.calls) expect(call).toEqual([]);
  });

  it('raises no banner when the cached account names a different tenant than the selection', async () => {
    // The tenant profile observed at boot is no evidence of a wrong-tenant
    // sign-in any more: a wrong-tenant token is impossible by construction
    // (getAccessToken asserts response.tenantId against the selection).
    const accountA = { homeAccountId: 'home.a', tenantId: TENANT_A.tenantId, username: 'user@example.test' };
    mockAccounts = [accountA];
    mockActiveAccount = accountA;

    await useAuthStore.getState().initializeMsal(TENANT_B);

    const { authError, accounts, isAuthenticated, msalInstance } = useAuthStore.getState();
    expect(authError).toBeNull();
    expect(isAuthenticated).toBe(true);
    expect(accounts).toEqual([accountA]);
    expect(msalInstance!.getActiveAccount()).toEqual(accountA);
  });

  it('adopts the first cached account when none is active', async () => {
    const accountA = { homeAccountId: 'home.a', tenantId: TENANT_A.tenantId, username: 'user@example.test' };
    mockAccounts = [accountA];
    mockActiveAccount = null;

    await useAuthStore.getState().initializeMsal(TENANT_A);

    const instance = useAuthStore.getState().msalInstance;
    expect(instance!.setActiveAccount).toHaveBeenCalledWith(accountA);
    expect(useAuthStore.getState().authError).toBeNull();
    expect(useAuthStore.getState().isAuthenticated).toBe(true);
  });

  it('leaves an existing active account alone', async () => {
    const accountA = { homeAccountId: 'home.a', tenantId: TENANT_A.tenantId, username: 'user@example.test' };
    mockActiveAccount = accountA;
    mockAccounts = [accountA];

    await useAuthStore.getState().initializeMsal(TENANT_A);

    const instance = useAuthStore.getState().msalInstance;
    expect(instance!.setActiveAccount).not.toHaveBeenCalled();
    expect(useAuthStore.getState().authError).toBeNull();
  });
});
