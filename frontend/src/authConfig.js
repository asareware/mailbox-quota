export const msalConfig = {
  auth: {
    clientId: "<FRONTEND_CLIENT_ID>",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost:8080"
  },
  cache: { cacheLocation: "sessionStorage" }
};

export const loginRequest = {
  scopes: ["api://<BACKEND_CLIENT_ID>/access_as_user", "openid", "profile", "offline_access"]
};
