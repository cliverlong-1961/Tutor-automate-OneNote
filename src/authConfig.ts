import type { Configuration } from "@azure/msal-browser";

export const msalConfig: Configuration = {
  auth: {
    clientId: "a8eef1c9-2398-453b-94b2-3efdbb2591f0",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + import.meta.env.BASE_URL,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const loginRequest = {
  scopes: [
    "User.Read",
    "Files.ReadWrite",
    "Notes.Create",
    "Notes.Read",
  ],
};
