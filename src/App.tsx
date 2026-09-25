import { useEffect } from 'react';
import { AppLayout } from '@/components/layout/AppLayout';
import { useAuthStore } from '@/stores/useAuthStore';
import { loadTenants, resolveTenant, readPersistedTenant, applyCanonicalTenantParam } from '@/lib/tenants';

function App() {
  const initializeMsal = useAuthStore((s) => s.initializeMsal);

  useEffect(() => {
    // Auto-init MSAL when returning from a redirect (code/error in URL)
    // or when an existing session is in sessionStorage (page reload after login).
    // Auth response may arrive in query string (?code=) or hash fragment (#code=)
    const params = new URLSearchParams(window.location.search);
    const hash = window.location.hash;
    const hasRedirectParams = params.has('code') || params.has('error')
      || hash.includes('code=') || hash.includes('error=');
    const hasMsalSession = Object.keys(sessionStorage).some((k) => k.startsWith('msal.'));

    if (hasRedirectParams || hasMsalSession) {
      // A genuine redirect return reads storage only (persistedId), never
      // the URL: navigateToLoginRequestUrl is false, so MSAL does not
      // restore the pre-redirect URL, and the tenant must never be re-derived
      // from window.location on this path. A load that merely finds
      // leftover msal.* keys (MSAL keeps them after a logout) has no pending
      // redirect to protect, so it honours ?tenant= like the normal leg.
      // loadTenants() never rejects — an empty list resolves to null, the
      // upstream default. Registered with the store synchronously so a Sign
      // in click that lands before the fetch resolves waits for this leg
      // (see login()).
      const boot = loadTenants().then(async (tenants) => {
        useAuthStore.getState().setTenants(tenants);
        const resolved = resolveTenant(tenants, {
          urlParam: hasRedirectParams ? null : params.get('tenant'),
          persistedId: readPersistedTenant(),
        });
        if (resolved && !hasRedirectParams && params.has('tenant')) {
          applyCanonicalTenantParam(resolved.tenantId);
        }
        // From here the automatic first switch is pending, and
        // initializeMsal is about to flip isAuthenticated — which is what
        // reveals the Switch tenant items. Mark the switch surface busy
        // first, so a click cannot land in the window between the session
        // appearing and the switch running and be dropped by the latch.
        if (resolved) useAuthStore.setState({ isSwitching: true });
        // Sign-in always starts at the default authority, so the instance
        // that consumes the pending response must be built for that same
        // authority; the tenant lives per request instead.
        await initializeMsal();
        return resolved;
      });

      // The boot promise settleBoot() waits on ENDS HERE, at initializeMsal.
      // The automatic switch below must stay outside it: switchTenant opens
      // with `await settleBoot()`, so a boot promise that contained the
      // switch could only settle after the switch it is blocking — a
      // permanent deadlock that also leaves the re-entrancy latch closed for
      // the rest of the page's life. Registered with the store synchronously
      // (before the fetch starts) so a Sign in click that lands early waits
      // for this leg (see login()).
      useAuthStore.getState().setBootPromise(boot.then(() => undefined));

      // The resolved tenant is not a picker preselection any more but the
      // session's automatic FIRST SWITCH, chained after boot has settled. No
      // confirmation dialog is needed — nothing is loaded yet, so the switch
      // clears empty stores and skips the reload (dataSource is 'none').
      // Run-once across React StrictMode's double invocation is supplied by
      // switchTenant's own re-entrancy latch and, from the second chain's
      // point of view afterwards, by the non-null selectedTenant.
      void boot
        .then((resolved) => {
          const { isAuthenticated, selectedTenant, switchTenant } = useAuthStore.getState();
          if (resolved && isAuthenticated && selectedTenant === null) {
            // switchTenant owns the flag from here: it releases it with its
            // own re-entrancy latch.
            return switchTenant(resolved);
          }
          useAuthStore.setState({ isSwitching: false });
        })
        .catch(() => {
          useAuthStore.setState({ isSwitching: false });
        });
    } else {
      // Normal boot (no redirect return, no existing session): populate the
      // list for the Switch tenant submenu and canonicalise ?tenant=<name> to
      // ?tenant=<guid> in the address bar. The rewrite happens only when the
      // user supplied a ?tenant= parameter: a choice that resolved from the
      // persisted browser preference alone must stay out of the address bar,
      // browser history and any URL the user copies. This leg is allowed to
      // read the URL — unlike the redirect-return leg above, nothing here
      // calls initializeMsal or sets selectedTenant, since normal boot stays
      // sample-mode-first per the module header comment in useAuthStore.ts.
      useAuthStore.getState().setBootPromise(
        loadTenants().then((tenants) => {
          useAuthStore.getState().setTenants(tenants);
          const resolved = resolveTenant(tenants, {
            urlParam: params.get('tenant'),
            persistedId: readPersistedTenant(),
          });
          if (resolved && params.has('tenant')) {
            applyCanonicalTenantParam(resolved.tenantId);
            // Nothing is persisted here. A deep link is a URL someone opened,
            // not a choice they made: writing ca-sim:tenant at boot let a
            // shared link silently and permanently rewrite the recipient's
            // remembered tenant, even if they closed the tab at once, and
            // their next unrelated visit would then auto-switch to it.
            // login() adopts the parameter instead — signing in is the act.
          }
        }),
      );
    }
  }, [initializeMsal]);

  return <AppLayout />;
}

export default App;
