// stores/useAuthStore.ts — Zustand store for deferred MSAL lifecycle.
//
// MSAL is NOT initialized at module scope. The app always mounts and renders
// sample mode. MSAL is created on-demand when the user clicks "Sign In" or
// when the app detects a redirect return (code/error in URL params). The
// authority the instance is built with is the page-load default only;
// selectedTenant is the tenant every token request is made against
// (services/auth.ts passes it per request), so it can change in place while
// the same instance and the same session stay live.
//
// Sign-in itself always happens at the DEFAULT authority — the account's own
// home tenant; a tenant is chosen only while connected, via Switch tenant.
// initializeMsal() keeps its `tenant` parameter even though no call site
// passes one any more: the per-request authority in services/auth.ts is what
// decides the tenant, so removing the parameter would delete that surface and
// its tests for no behavioural gain. What the parameter must NOT do is reset
// the selection; see the guarded write in initializeMsal.

import { create } from 'zustand';
import {
  PublicClientApplication,
  EventType,
} from '@azure/msal-browser';
import type { EventMessage, AuthenticationResult, AccountInfo } from '@azure/msal-browser';
import { buildMsalConfig, loginRequest } from '../authConfig';
import { authorityFor, persistTenant, resolveTenant } from '../lib/tenants';
import type { TenantEntry } from '../lib/tenants';
import { usePolicyStore } from './usePolicyStore';
import { usePersonaStore } from './usePersonaStore';
import { useEvaluationStore } from './useEvaluationStore';

interface AuthState {
  msalInstance: PublicClientApplication | null;
  isInitialized: boolean;
  isInitializing: boolean;
  authError: string | null;
  accounts: AccountInfo[];
  isAuthenticated: boolean;
  selectedTenant: TenantEntry | null;
  // The tenant a fresh, interactive sign-in would cure — set together with
  // the refusal banner text when the selected tenant refuses a silent token,
  // and the only thing that makes the banner's "Sign in to X" button appear.
  // Never a standing selection: cleared the moment the banner is dismissed,
  // a switch happens, or the re-authentication starts.
  reauthTenant: TenantEntry | null;
  // True while a switch owns the critical section below, and — set by the
  // boot leg — while the session's automatic first switch is still pending.
  // The Switch tenant items read it: isAuthenticated flips inside
  // initializeMsal, i.e. BEFORE the automatic switch runs, so without this a
  // click landing in that window is dropped by the re-entrancy latch with no
  // state change, no error and no feedback.
  isSwitching: boolean;
  // The stable, realm-independent identifier of the signed-in account — the
  // ONLY account key that survives MSAL's one-profile behaviour, where the
  // account entity keeps a tenant profile for the most recently minted token
  // alone. It is session state: logout() clears it, switchTenant() must not.
  homeAccountId: string | null;
  // Configuration read from public/tenants.json at boot, NOT session state —
  // never cleared by logout()/switchTenant(), which only reset the live
  // MSAL/account session. Clearing it there would blank the picker after a
  // switch even though the operator-curated list itself hasn't changed.
  tenants: TenantEntry[];

  // Actions
  initializeMsal: (tenant?: TenantEntry | null) => Promise<void>;
  login: () => Promise<void>;
  logout: () => Promise<void>;
  switchTenant: (tenant: TenantEntry) => Promise<void>;
  reauthenticate: (tenant: TenantEntry) => Promise<void>;
  clearAuthError: () => void;
  setTenants: (tenants: TenantEntry[]) => void;
  setBootPromise: (promise: Promise<void> | null) => void;
}

// Single in-flight init promise to prevent concurrent initialization
let initPromise: Promise<void> | null = null;

// The app's boot effect (tenants.json fetch, then — on a redirect return —
// initializeMsal). Registered by App.tsx before the fetch starts so that
// login()/switchTenant() can settle it: between mount and the fetch
// resolving, tenants is [] and the plain Sign in button is what renders, and
// a click there must not build a default-authority instance that consumes
// the pending redirect and drops the persisted tenant for this page load.
let bootPromise: Promise<void> | null = null;

async function settleBoot(): Promise<void> {
  // Neither leg of the boot effect rejects today; the catch keeps a future
  // rejection from wedging every later sign-in attempt.
  if (bootPromise) await bootPromise.catch(() => {});
  if (initPromise) await initPromise;
}

/**
 * The `?tenant=` parameter of the page the user is signing in FROM, resolved
 * against the curated list (never a synthesized entry — see resolveTenant).
 * Read here rather than handed in by App.tsx so nothing about a deep link is
 * stored before it is acted on. App.tsx has already canonicalised the
 * parameter to the tenant id by the time any click can reach this.
 */
function readUrlTenant(tenants: TenantEntry[]): TenantEntry | null {
  if (typeof window === 'undefined' || !window.location) return null;
  const urlParam = new URLSearchParams(window.location.search ?? '').get('tenant');
  return resolveTenant(tenants, { urlParam });
}

// Re-entrancy latch for switchTenant, guarding the CRITICAL SECTION only —
// persist, re-point, clear. A second click in that window is dropped rather
// than interleaving two clear sequences. The data reload that follows is
// deliberately outside it: holding the latch across a multi-second Graph
// load would drop the next switch silently, with no state change and no
// feedback, leaving the user believing they had switched.
let switchInProgress = false;

// Which switch owns the reload. Only the newest one runs it; usePolicyStore's
// own load generation discards any older load that is still in flight.
let switchGeneration = 0;

export const useAuthStore = create<AuthState>((set, get) => ({
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

  initializeMsal: async (tenantArg?: TenantEntry | null) => {
    const tenant = tenantArg ?? null;
    const { isInitialized, isInitializing } = get();
    if (isInitialized || isInitializing) {
      // If already initializing, wait for the in-flight promise
      if (initPromise) await initPromise;
      return;
    }

    const clientId = import.meta.env.VITE_MSAL_CLIENT_ID;
    if (!clientId) {
      set({
        authError:
          'Missing MSAL configuration. Set VITE_MSAL_CLIENT_ID in a .env file to connect to a live tenant. See README.',
      });
      return;
    }

    // `tenant ?? get().selectedTenant`, never a bare `tenant`: production
    // calls this with no argument (App.tsx) or with null (login()), so an
    // unconditional write would silently drop a selection made by
    // switchTenant and send every later token request to the home tenant with
    // nothing on screen changing. The parameter is only ever a STARTING
    // selection for an instance being built, never a reset.
    set({
      isInitializing: true,
      authError: null,
      reauthTenant: null,
      selectedTenant: tenant ?? get().selectedTenant,
    });

    initPromise = (async () => {
      try {
        const instance = new PublicClientApplication(buildMsalConfig(authorityFor(tenant)));
        await instance.initialize();

        // Register event callback BEFORE handleRedirectPromise to avoid
        // missing LOGIN_SUCCESS (same ordering as the previous main.tsx).
        // Unfiltered, like the post-redirect write below, so both writes
        // agree on what `accounts` means.
        instance.addEventCallback((event: EventMessage) => {
          if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
            const payload = event.payload as AuthenticationResult;
            instance.setActiveAccount(payload.account);
            set({
              accounts: instance.getAllAccounts(),
              homeAccountId: payload.account?.homeAccountId ?? null,
              isAuthenticated: true,
            });
          }
        });

        // Process any pending redirect response
        await instance.handleRedirectPromise();

        // Unfiltered account resolution — byte-for-byte the upstream
        // behaviour, for every tenant selection. A tenant filter here would
        // answer empty for the tenant switched away from: MSAL keeps a tenant
        // profile only for the most recently minted token, so the filtered
        // list is no statement about the session at all.
        const accounts = instance.getAllAccounts();

        // Upstream's rule: adopt the first cached account when none is
        // active (e.g. page refresh).
        if (!instance.getActiveAccount() && accounts.length > 0) {
          instance.setActiveAccount(accounts[0]);
        }

        // No boot-time tenant-mismatch banner: a wrong-tenant token is
        // impossible by construction now — getAccessToken asserts
        // response.tenantId against the selection — so the refusal banner
        // belongs to token acquisition, not to a tenant profile observed
        // here, which names whichever tenant happened to get a token last.
        set({
          msalInstance: instance,
          isInitialized: true,
          accounts,
          isAuthenticated: accounts.length > 0,
          // The account MSAL considers ACTIVE, never an arbitrary cache
          // entry. handleRedirectPromise() above is what fires LOGIN_SUCCESS,
          // and that callback has already made the account that just signed
          // in the active one — so this write must not clobber it with
          // getAllAccounts()[0], which can be a different user left behind by
          // an abandoned logout (MSAL keeps msal.* keys). getAccessToken
          // resolves the token account from this id alone, and its response
          // assertion cannot catch a wrong-user token: the tenant is right,
          // the identity is not. The fallback is reached only when nothing is
          // active, i.e. the adopt-first-cached-account rule above found no
          // accounts at all.
          homeAccountId: (instance.getActiveAccount() ?? accounts[0])?.homeAccountId ?? null,
          authError: null,
        });
      } catch {
        // Fixed text only: the authority now embeds the tenant id, and MSAL's
        // endpoint-resolution and network errors quote the authority URL in
        // their message — which must reach neither the banner nor the console
        // from here. (MSAL's own logger relay in authConfig.ts redacts the
        // tenant id from the lines it forwards.)
        console.error('MSAL initialization failed');
        set({
          authError: 'Could not initialize sign-in for the selected tenant. Check your network and try again.',
        });
      } finally {
        set({ isInitializing: false });
        initPromise = null;
      }
    })();

    await initPromise;
  },

  login: async () => {
    // The boot effect may still be fetching tenants.json or building an
    // instance (MSAL initialize() + handleRedirectPromise() pending). Settle
    // that first — otherwise initializeMsal()'s isInitializing early-return
    // is reached before the boot leg has consumed a pending redirect, or a
    // click before the fetch resolves builds an instance ahead of that leg.
    await settleBoot();

    // Sign-in always targets the account's home tenant via the default
    // authority; nothing is chosen (or persisted) here. A remembered or
    // deep-linked tenant is applied afterwards, as the session's automatic
    // first switch (App.tsx).
    if (get().isInitialized && get().isAuthenticated) {
      // The boot leg just finished a redirect return (the click landed while
      // it was in flight). Nothing to sign in to.
      return;
    }

    if (get().isInitialized && get().selectedTenant !== null) {
      // Signed out, but an instance was already built for a tenant authority
      // (MSAL leaves msal.* keys in sessionStorage after a logout, so the
      // boot effect fires on every later page load). Nothing is in flight on
      // an idle instance, so discard it and rebuild at the default authority.
      //
      // Unreachable by construction TODAY: no production call site passes a
      // tenant to initializeMsal, and a non-null selectedTenant otherwise
      // requires switchTenant, which requires an authenticated session — the
      // early return above. It stays because it is the guarantee for the
      // parameter that stays with it: if anything ever initialises for a
      // tenant authority again, the plain Sign in button must still rebuild
      // at the default one rather than sign in to that tenant. Deleting it
      // would also delete the only coverage of that contract.
      set({ msalInstance: null, isInitialized: false, accounts: [], selectedTenant: null });
    }

    await get().initializeMsal(null);
    const instance = get().msalInstance;
    if (!instance) return; // initializeMsal failed — authError already set

    // Adopting a deep-linked tenant is a USER act, and this is it. Persisting
    // it at boot instead would let a shared link rewrite the remembered
    // tenant of anyone who merely opened it. It must happen before the
    // redirect (which unloads the page) because the return leg reads storage
    // only: this is the channel by which a deep-linked tenant
    // survives to become the automatic first switch. A tenant that resolved
    // from the persisted value alone is already stored and is not rewritten.
    const urlTenant = readUrlTenant(get().tenants);
    if (urlTenant) persistTenant(urlTenant.tenantId);

    await instance.loginRedirect(loginRequest);
  },

  logout: async () => {
    const instance = get().msalInstance;
    if (!instance) return;
    // Clear all stores before redirect (prevents stale tenant data)
    usePolicyStore.getState().clear();
    usePersonaStore.getState().clear();
    useEvaluationStore.getState().clear();
    // Session state, cleared with the session it describes.
    set({ homeAccountId: null, reauthTenant: null });
    await instance.logoutRedirect();
  },

  switchTenant: async (tenant: TenantEntry) => {
    // Clear, re-point, reload — one session, one token per tenant. Nothing
    // navigates, so a second invocation is only dropped while the first is
    // still inside the critical section (the latch is released in the finally
    // below, before the reload).
    if (switchInProgress) return;
    switchInProgress = true;
    set({ isSwitching: true });
    const myGeneration = ++switchGeneration;
    let reload = false;

    try {
      // Settle an in-flight boot initialization first — otherwise a switch
      // that arrives while msalInstance is still null decides whether to
      // reload before boot has produced an instance at all.
      await settleBoot();

      // Read before the stores are cleared: only a tenant whose data was on
      // screen is reloaded, so a switch in sample mode or before any load
      // keeps the existing "Load from Tenant" affordance as the next step.
      const wasLive = usePolicyStore.getState().dataSource === 'live';

      // Persist first: this is the choice the next page load must resolve.
      persistTenant(tenant.tenantId);
      set({ selectedTenant: tenant, authError: null, reauthTenant: null });

      // Clear all stores exactly as logout() does — the previous tenant's
      // policies, personas and evaluation results must not survive a switch.
      usePolicyStore.getState().clear();
      usePersonaStore.getState().clear();
      useEvaluationStore.getState().clear();

      const instance = get().msalInstance;
      if (!instance) {
        // Sample mode or never signed in — there is no session to re-point.
        // The persisted choice above is enough for the next login(tenant).
        return;
      }

      reload = wasLive;
    } finally {
      switchInProgress = false;
      // Released with the latch, not after the reload: the reload is outside
      // the critical section precisely so the next switch is accepted.
      set({ isSwitching: false });
    }

    // The session stays signed in; the new tenant's token is acquired against
    // its own authority on the next request (services/auth.ts). This reload is
    // what makes the new tenant's policies appear without a further click. It
    // runs outside the latch, so a switch requested while it is in flight is
    // accepted — and last-write-wins: that newer switch bumps the generation
    // and owns the reload from here on. The rejection is swallowed on purpose:
    // the policy store records its own error and getAccessToken raises the
    // auth banner when the new tenant demands interaction. It is not logged
    // either — that error text can carry the authority URL.
    if (reload && myGeneration === switchGeneration) {
      await usePolicyStore.getState().loadFromGraph().catch(() => {});
    }
  },

  reauthenticate: async (tenant: TenantEntry) => {
    // The cure for a tenant's token refusal, and the ONLY interactive flow the
    // app starts for a selected tenant — never automatically: it runs from the
    // banner's button, i.e. a user act.
    //
    // prompt: 'login' is load-bearing. The usual refusal is the target
    // tenant's sign-in frequency policy rejecting the silent redeem on
    // auth_time (AADSTS70043); a sign-in that rides
    // the existing SSO session inherits the old auth_time and would be
    // refused again, so only a genuine re-authentication cures it.
    //
    // Redirect, not popup (the app's only interactive shape: a popup loads
    // the whole SPA and never closes). Nothing is persisted or cleared here:
    // the remembered tenant stays, so after the return the boot leg's
    // automatic first switch lands back in this tenant — now with a fresh
    // auth_time.
    const instance = get().msalInstance;
    if (!instance) return;
    set({ authError: null, reauthTenant: null });
    await instance.acquireTokenRedirect({
      ...loginRequest,
      authority: authorityFor(tenant),
      prompt: 'login',
    });
  },

  clearAuthError: () => {
    set({ authError: null, reauthTenant: null });
  },

  setTenants: (tenants: TenantEntry[]) => {
    set({ tenants });
  },

  setBootPromise: (promise: Promise<void> | null) => {
    bootPromise = promise;
  },
}));
