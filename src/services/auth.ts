// services/auth.ts — Centralized token acquisition.
//
// The MSAL instance is managed by useAuthStore (deferred initialization).
// This module exports getAccessToken() which reads the instance from the store.

import { loginRequest } from '../authConfig';
import { useAuthStore } from '../stores/useAuthStore';
import { isSilentRenewalFailure } from '../lib/authErrors';
import { authorityFor, sameTenant } from '../lib/tenants';

// Concurrent getAccessToken calls (e.g. policy load + persona resolution) can
// both hit a renewal failure; a second acquireTokenRedirect while the first is
// starting throws interaction_in_progress. Only the first caller triggers the
// redirect.
let redirectInProgress = false;

/**
 * Thrown when a token comes back for a tenant other than the selected one —
 * the load-bearing cross-tenant check. Distinct from the "no account in the
 * cache" rejection: the recovery is not "sign in", and a shared message would
 * make the two indistinguishable in logs, in tests and while debugging. Fixed
 * text, no tenant id.
 */
export const TENANT_MISMATCH_ERROR = 'Token issued for a different tenant than the one selected';

/**
 * Acquire an access token silently using the active MSAL account.
 * If the silent call fails due to expired session or missing consent,
 * automatically triggers a redirect-based login flow.
 */
export async function getAccessToken(): Promise<string> {
  const msalInstance = useAuthStore.getState().msalInstance;
  if (!msalInstance) throw new Error('MSAL not initialized');

  const { selectedTenant, homeAccountId } = useAuthStore.getState();

  let account;

  if (!selectedTenant) {
    // Upstream path, byte-identical: the active account, else the first
    // cached one (e.g. page refresh), else nothing.
    account = msalInstance.getActiveAccount();
    if (!account) {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        account = accounts[0];
      }
    }
  } else {
    // Resolve the signed-in account by home account id ONLY. MSAL keeps a
    // single tenant profile on the account entity — the tenant of the most
    // recently minted token — so after a switch
    // both getActiveAccount() (its filter names the profile that vanished)
    // and getAllAccounts({ tenantId }) answer nothing for the tenant switched
    // away from. getAllAccounts({ homeAccountId }) does not go through tenant
    // profiles and is the one lookup that survives.
    //
    // There is no fallback when homeAccountId is null: guessing from an
    // unfiltered getAllAccounts()[0] could pick a different cached user left
    // over from an abandoned logout, and the response assertion below cannot
    // catch that (same tenant, other user). It fails closed instead.
    const base = homeAccountId ? msalInstance.getAllAccounts({ homeAccountId })[0] : undefined;
    if (base) {
      // MSAL's silent flow takes its tenant context from
      // request.account.tenantId, so stamping the selection here is what makes
      // it look for that realm's cached access token — or redeem the
      // realm-independent refresh token (keyed by homeAccountId) when there is
      // none. The active account is deliberately NOT re-pointed: it keeps
      // describing the account that is actually signed in.
      account = { ...base, tenantId: selectedTenant.tenantId };
    }
  }
  if (!account) throw new Error('No active account');

  // The instance's configured authority is only the page-load default; the
  // tenant a token is issued for is decided per request, here. authorityFor()
  // is the single place a tenant id becomes an authority URL (it re-validates
  // the GUID). With no tenant selected the request is the upstream one.
  const request = selectedTenant
    ? { ...loginRequest, account, authority: authorityFor(selectedTenant) }
    : { ...loginRequest, account };

  let response;
  try {
    response = await msalInstance.acquireTokenSilent(request);
  } catch (error) {
    // Interaction-required and a timed-out hidden iframe are the same situation
    // with and without an answer from Entra: the session is gone, so send the
    // user to sign in rather than surfacing an MSAL error code.
    if (isSilentRenewalFailure(error)) {
      if (selectedTenant) {
        // The selected tenant refuses a silent token (sign-in frequency,
        // missing consent, missing MFA trust). The app never starts a login
        // flow by itself for a tenant the user selected — the banner offers
        // it, and reauthTenant is what makes that button appear. Signing out
        // is NOT the cure and is no longer suggested: sign-in always targets
        // the home tenant, so it would return here and be refused again.
        // Fixed text, tenant named by its configured name only: MSAL error
        // text and the authority URL both carry the tenant id.
        useAuthStore.setState({
          authError: `Could not get access to ${selectedTenant.name} without signing in again.`,
          reauthTenant: selectedTenant,
        });
        throw new Error('Interactive sign-in required for the selected tenant');
      }
      if (!redirectInProgress) {
        redirectInProgress = true;
        await msalInstance.acquireTokenRedirect(loginRequest);
      }
      // acquireTokenRedirect navigates away — this line is reached only by
      // concurrent callers while the redirect is starting.
      throw new Error('Redirecting to login');
    }
    if (selectedTenant) {
      // Everything else a cross-tenant redeem can answer — the app's service
      // principal missing in the target tenant, a server error, a rejected
      // refresh token — arrives as MSAL text that quotes the authority URL,
      // the tenant id and correlation ids. The policy store stores this
      // message verbatim and the sidebar renders it, so the MSAL text must
      // not leave this function. Fixed text, tenant named by its configured
      // name only.
      throw new Error(
        `Could not get access to ${selectedTenant.name}. Check that this app is available in that tenant, then try again.`,
      );
    }
    throw error;
  }

  // The load-bearing cross-tenant guard, outside the try so its own rejection
  // is not remapped by the catch above: a cache hit for the wrong tenant is
  // the last way a tenant-A token could reach a tenant-B Graph call.
  //
  // response.tenantId alone, and no fallback to the account: it is
  // non-optional in msal-browser 5 and always populated for a successful
  // redeem, so a fallback could only ever fire on the malformed response this
  // guard exists for — and request.account's tenant is the selection itself,
  // stamped on above, so comparing against it would wave every wrong-tenant
  // response through. A guard whose degraded mode is "accept" is the wrong
  // default here.
  if (selectedTenant && !sameTenant(response.tenantId, selectedTenant.tenantId)) {
    // One distinguishable exception, and only this one: a CACHE HIT whose
    // realm ID token MSAL dropped. CacheManager.getIdToken deletes the
    // candidates and returns null when it finds more than one ID token for
    // the realm, so response.tenantId is "" while that realm's access token
    // is returned intact. The lookup was scoped by request.account.tenantId —
    // the selection, stamped on above — so the token is the selected realm's
    // by construction, and reporting it as a cross-tenant token would name a
    // security event where there is none and send the next debugger down the
    // wrong path. On the network path an absent tid is unexplained, so it
    // still fails closed; so does any NON-empty mismatching tenantId,
    // cache hit or not.
    if (!(response.fromCache === true && !response.tenantId)) {
      throw new Error(TENANT_MISMATCH_ERROR);
    }
  }
  return response.accessToken;
}
