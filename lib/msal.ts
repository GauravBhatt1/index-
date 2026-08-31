import { PublicClientApplication, type Configuration } from '@azure/msal-browser';
import { PUBLIC_AZURE_CLIENT_ID, PUBLIC_AZURE_TENANT_ID } from '$env/static/public';

const clientId = PUBLIC_AZURE_CLIENT_ID || '';
const tenant = PUBLIC_AZURE_TENANT_ID || 'common';

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
