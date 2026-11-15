import React from "react";
import ReactDOM from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";

import App from "./App";
import { msalConfig } from "./authConfig";
import "./index.css";

const msalInstance = new PublicClientApplication(msalConfig);

async function startApp() {
  await msalInstance.initialize();

  // Process the response returned by Microsoft after sign-in.
  const response = await msalInstance.handleRedirectPromise();

  // Set the authenticated account as the active account.
  if (response?.account) {
    msalInstance.setActiveAccount(response.account);
  } else {
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
    }
  }

  ReactDOM.createRoot(
    document.getElementById("root")!
  ).render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </React.StrictMode>
  );
}

startApp().catch(console.error);