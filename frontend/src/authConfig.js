// Fill these placeholders
export const FRONTEND_CLIENT_ID = "4927f6c4-fada-4ba3-b71a-ca1af2c7ea24";
export const TENANT_ID = "81662b37-766f-4e70-82f0-bf7eae560494";
export const BACKEND_CLIENT_ID = "fb30286b-66d4-499e-ac48-0b579e070444";

export const msalConfig = {
  auth: {
    clientId: FRONTEND_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: "http://localhost:8080"
  },
  cache: { cacheLocation: "sessionStorage" }
};

export const loginRequest = {
  scopes: [
    `api://${BACKEND_CLIENT_ID}/access_as_user`,
    "openid",
    "profile",
    "offline_access"
  ]
};
