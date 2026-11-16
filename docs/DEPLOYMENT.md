# Deployment

## Components

The known deployment endpoints are:

- Frontend: `https://tutor-video-onenote.crlong.uk`
- Backend (Azure Container Apps): `https://tutoring-automation-api.graycliff-b3a723e4.uksouth.azurecontainerapps.io`

The exact automation and deployment triggers have not yet been verified from the repository and Azure configuration.

## Frontend deployment checks

Before publishing, verify:

1. Vite `base` setting in `vite.config.ts`.
2. Build command and output directory in `package.json`.
3. GitHub Actions workflow files under `.github/workflows/`.
4. GitHub Pages publishing source and custom-domain configuration.
5. MSAL redirect URI registered for the deployed frontend.
6. Production frontend origin is allowed by backend `CORS_ORIGINS`.

Do not assume a push to a particular branch automatically publishes the site until the workflow has been inspected.

## Backend deployment checks

The backend runs in Azure Container Apps. A source-code push does not necessarily update the running container. Verify whether deployment is triggered by GitHub Actions, an image registry build, a manual process or another mechanism.

Ensure the deployed container has the required dependencies and environment configuration. Set API keys using Azure's supported secret configuration; do not bake them into the image or commit them.

## Post-deployment test

1. Open the deployed frontend and sign in.
2. Check backend `/health`.
3. Check AssemblyAI and Gemini diagnostics.
4. Browse a test OneDrive folder.
5. Process a short non-confidential recording.
6. Confirm transcript and summary uploads.
7. Generate a PDF.
8. Create a OneNote page in a test notebook.
9. Check browser network errors and Azure logs.

## Details still to establish

See [Open questions](OPEN-QUESTIONS.md) for Azure resource names, registry, deployment trigger, GitHub Pages workflow, environment settings and production CORS configuration.
