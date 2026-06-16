export const msalConfig = {
  auth: {
    clientId: "7269d452-36d7-4686-be65-d343c65ad02b",
    authority: "https://login.microsoftonline.com/7163abb0-3c0a-412b-b0c8-c9fdda5b1d3b",
    redirectUri: "https://thankful-pond-052fbeb0f.7.azurestaticapps.net",
    postLogoutRedirectUri: "https://thankful-pond-052fbeb0f.7.azurestaticapps.net",
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};

export const loginRequest = {
  scopes: [
    "https://analysis.windows.net/powerbi/api/Dataset.Read.All",
    "https://analysis.windows.net/powerbi/api/Workspace.Read.All"
  ],
  redirectUri: "https://thankful-pond-052fbeb0f.7.azurestaticapps.net",
  responseMode: "query"
};