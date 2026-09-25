// authConfig.ts
 
// MSAL configuration for silent token acquisition and caching.
// Reference: https://learn.microsoft.com/en-us/entra/identity-platform/msal-acquire-cache-tokens

import { LogLevel } from '@azure/msal-browser';
import { DEFAULT_AUTHORITY } from './lib/tenants';

// lib/guid.ts anchors its copy to the whole string (it answers "is this value
// a GUID?"); this one scans a free-text log line for GUIDs anywhere in it.
const GUID_RE = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi;

/**
 * Builds the MSAL config for a given authority. Every non-authority field is
 * fixed; only `auth.authority` varies per invocation, so a per-tenant
 * PublicClientApplication can be constructed without duplicating the rest of
 * this config at each call site. Defaults to the upstream `organizations`
 * authority when no tenant has been chosen.
 */
export function buildMsalConfig(authority: string = DEFAULT_AUTHORITY) {
  // MSAL's endpoint-resolution and network errors quote the authority URL,
  // which embeds a tenant id. The needle cannot be derived from `authority`:
  // this instance is built for the page-load tenant but serves every tenant
  // the user switches to, because services/auth.ts sets the authority per
  // request. So redact by shape rather than by value — every GUID in a
  // relayed line goes, whichever tenant raised it. Lines that carry no GUID
  // (the default-authority case among them) are relayed unchanged.
  const redact = (message: string) => message.replace(GUID_RE, '<tenant>');
  return {
    auth: {
      clientId: import.meta.env.VITE_MSAL_CLIENT_ID,
      authority,
      redirectUri: window.location.origin,
      postLogoutRedirectUri: '/',
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: 'sessionStorage' as const,
      storeAuthStateInCookie: false,
    },
    system: {
      loggerOptions: {
        loggerCallback: (level: LogLevel, message: string, containsPii: boolean) => {
          if (containsPii) return;
          if (level === LogLevel.Error) console.error(redact(message));
          else if (level === LogLevel.Warning) console.warn(redact(message));
        },
      },
    },
  };
}

/**
 * Scopes for Graph API access.
 * MSAL.js automatically adds OIDC scopes (openid, profile, email).
 *
 * Directory.Read.All covers users, groups, directory roles, role templates,
 * service principals, and applications — so the previously requested
 * User.Read.All / GroupMember.Read.All / Application.Read.All were redundant
 * consent surface. Policy.Read.All is still needed for CA policies, named
 * locations, and authentication strength policies.
 */
export const loginRequest = {
  scopes: [
    'Policy.Read.All',
    'Directory.Read.All',
  ],
};
