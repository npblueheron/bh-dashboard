// src/services/authConfig.js
export const msalConfig = {
  auth: {
    clientId: "434f0a49-7cb8-4805-a62d-65a385169920",
    authority: "https://login.microsoftonline.com/7163abb0-3c0a-412b-b0c8-c9fdda5b1d3b",
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

export const loginRequest = {
  scopes: ["https://analysis.windows.net/powerbi/api/Dataset.Read.All"]
};

// Power BI workspace and dataset
export const PBI_CONFIG = {
  workspaceId: null, // will be discovered
  datasetId: null,   // will be discovered
  workspaceName: "WS_BH_DATA_DEV",
  datasetName: "Project_Dashboard_Model"
};
