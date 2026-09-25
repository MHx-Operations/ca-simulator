/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_MSAL_CLIENT_ID: string;
  readonly VITE_TENANTS_URL?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
