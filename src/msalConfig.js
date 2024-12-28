import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "4bd54a16-f4f5-466b-9309-276f9850a53f", 
    authority: "https://login.microsoftonline.com/eb959c4a-a5cf-47c0-bf88-2cc75d358cf3", 
    redirectUri: "http://localhost:3000", // https://d2gggx7ht3ipka.cloudfront.net http://localhost:3000 
  },
  cache: {
    cacheLocation: "localStorage", // O "sessionStorage"
    storeAuthStateInCookie: false,
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

export default msalInstance;
