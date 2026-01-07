import { Configuration, LogLevel } from "@azure/msal-browser";

const tenantId = import.meta.env.VITE_TENANT_ID;
const redirectUri = import.meta.env.VITE_REDIRECT_URI || window.location.origin;

const defaultScopes = ["User.Read", "email", "openid", "profile", "Sites.ReadWrite.All"];

export const loginRequest = {
  scopes: (import.meta.env.VITE_GRAPH_SCOPES || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean).length
    ? (import.meta.env.VITE_GRAPH_SCOPES || "").split(",").map((s) => s.trim())
    : defaultScopes,
};

export const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_CLIENT_ID,
    authority: tenantId ? `https://login.microsoftonline.com/${tenantId}` : "https://login.microsoftonline.com/common",
    redirectUri,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Error,
      loggerCallback: () => {},
    },
  },
};

export const protectedResources = {
  sharePointSiteId: import.meta.env.VITE_SHAREPOINT_SITE_ID,
  sharePointListId: import.meta.env.VITE_SHAREPOINT_LIST_ID,
  powerAutomateFlowUrl: import.meta.env.VITE_PA_FLOW_URL,
};
