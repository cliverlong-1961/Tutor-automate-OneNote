# Troubleshooting

## Frontend cannot reach backend

Check the `API_BASE` value in `src/App.tsx`, Azure Container Apps ingress/availability, and backend `CORS_ORIGINS`. Inspect browser developer tools for network and CORS errors.

## Microsoft Graph HTTP 401

Some frontend operations refresh the token and retry once. If the operation still fails:

1. Sign out and sign in again.
2. Confirm the signed-in user can access the OneDrive folder or OneNote section.
3. Check the app registration redirect URI and delegated Graph permissions.
4. Check whether consent is required.

Never share the access token.

## AssemblyAI transcription failure

- Check `/health` for AssemblyAI configuration status.
- Use `/diagnostics/assemblyai`.
- Confirm the file is MP4 or MKV.
- Check OneDrive download access and backend logs.
- For MKV, confirm FFmpeg is installed and `FFMPEG_PATH` is correct.

## Gemini summary failure

- Check `/health` for Gemini configuration.
- Use `/test-gemini`.
- Check the Gemini key, model and timeout settings.
- Inspect backend logs for authentication, quota or timeout errors.

The backend saves the transcript before requesting the summary. If Gemini fails, use **Generate summary from existing transcript** rather than transcribing again.

## PDF generation failure

Use `/diagnostics/pdf`; check backend logs for rendering, font or upload errors. Confirm the summary is non-empty and that the backend can upload to OneDrive.

## OneNote page creation failure

Confirm a notebook and section are selected, the summary is loaded, preparation has created the PDF and sharing links, the PDF remains in OneDrive, and the user can create pages in the selected section. Inspect the Graph error response without disclosing tokens or lesson content.

## Incorrect lesson status

File matching is based on filename stems within the currently loaded OneDrive folder. Confirm that recording, transcript and summary names match. The inspected lesson grouping filters for MP4, although the processing picker lists MP4 and MKV.

## Local backend setup

Confirm Python is installed, `.venv` exists, requirements install successfully, `backend/.env` is configured and FFmpeg is available if MKV processing is required. Do not delete `.env` without securely preserving its configuration.
