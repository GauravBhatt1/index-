import { PublicClientApplication, type Configuration } from '@azure/msal-browser';

const clientId = import.meta.env.NEXT_PUBLIC_AZURE_CLIENT_ID || '';
const tenant = import.meta.env.NEXT_PUBLIC_AZURE_TENANT_ID || 'common';

export const graphScopes = ['Files.Read'];

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenant}`,
    redirectUri: typeof window !== 'undefined' ? window.location.origin : undefined,
    postLogoutRedirectUri: typeof window !== 'undefined' ? window.location.origin : undefined
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false
  }
};

let instance: PublicClientApplication | null = null;
export function getMsal() {
  if (!instance) instance = new PublicClientApplication(msalConfig);
  return instance;
}
