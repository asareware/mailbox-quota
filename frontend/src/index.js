// frontend/src/index.js
import React from "react";
import { createRoot } from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import App from "./App";
import { msalConfig } from "./authConfig";

const msalInstance = new PublicClientApplication(msalConfig);

const rootElement = document.getElementById("root");
if (!rootElement) {
  // Defensive: helpful error message if index.html missing <div id="root">
  throw new Error("Root element not found. Ensure public/index.html contains <div id=\"root\"></div>.");
}

const root = createRoot(rootElement);
root.render(
  <MsalProvider instance={msalInstance}>
    <App />
  </MsalProvider>
);