// set this to the origin(s) you allow; can be from env if you prefer
const ALLOWED_ORIGINS = (process.env.CORS_ORIGINS && process.env.CORS_ORIGINS.split(',')) || ['http://localhost:3000'];

function allowedOriginForRequest(reqOrigin) {
  if (!reqOrigin) return null;
  if (ALLOWED_ORIGINS.includes('*')) return '*';
  return ALLOWED_ORIGINS.includes(reqOrigin) ? reqOrigin : null;
}

/** helper to attach CORS headers to a response */
function corsHeadersForOrigin(origin) {
  return {
    'Access-Control-Allow-Origin': origin || '',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Access-Control-Allow-Credentials': 'true'
  };
}

module.exports = async function (context, req) {
  context.log('MailFoldersFunction (multi-tenant) triggered');

  const requestOrigin = req.headers?.origin;
  const allowedOrigin = allowedOriginForRequest(requestOrigin);

  // If it's a preflight request, respond quickly:
  if (req.method === 'OPTIONS') {
    context.res = {
      status: 204,
      headers: corsHeadersForOrigin(allowedOrigin)
    };
    return;
  }

  try {
    // ... normal token checks and logic follow

    // At the end before returning success:
    context.res = {
      status: 200,
      headers: Object.assign({ 'Content-Type': 'application/json' }, corsHeadersForOrigin(allowedOrigin)),
      body: response
    };
  } catch (err) {
    // include CORS headers on errors too
    context.res = {
      status: err.status || 500,
      headers: corsHeadersForOrigin(allowedOrigin),
      body: { error: err.message || 'Server error' }
    };
  }
};


import React from "react";
import { createRoot } from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import App from "./App";
import { msalConfig } from "./authConfig";

const msalInstance = new PublicClientApplication(msalConfig);
const root = createRoot(document.getElementById("root"));
root.render(
  <MsalProvider instance={msalInstance}>
    <App />
  </MsalProvider>
);