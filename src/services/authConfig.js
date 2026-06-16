export const msalConfig = {
  auth: {
    clientId: "7269d452-36d7-4686-be65-d343c65ad02b",
    authority: "https://login.microsoftonline.com/7163abb0-3c0a-412b-b0c8-c9fdda5b1d3b",
    redirectUri: "https://dashboard.blueheron.com",
    postLogoutRedirectUri: "https://dashboard.blueheron.com"
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
  ]
};