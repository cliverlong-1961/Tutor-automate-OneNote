# Architecture

## Overview

```text
Browser: React + TypeScript + Vite
  |-- MSAL sign-in and delegated Microsoft Graph token
  |-- Microsoft Graph: profile, OneDrive, sharing links, OneNote
  |
  +--> FastAPI backend (Azure Container Apps)
         |--> Microsoft Graph / OneDrive download and upload
         |--> FFmpeg (MKV audio extraction)
         |--> AssemblyAI (transcription)
         |--> Google Gemini (revision summary)
         |--> PDF rendering libraries
```

## Frontend

Key files inspected:

- `src/App.tsx` — main application state and workflow.
- `src/authConfig.ts` — MSAL configuration and Graph scopes.
- `src/main.tsx` — application entry point.
- `src/App.css` and `src/index.css` — styling.

The configured backend URL in `App.tsx` is:

`https://tutoring-automation-api.graycliff-b3a723e4.uksouth.azurecontainerapps.io`

MSAL uses the Microsoft identity platform, `sessionStorage`, and delegated scopes including `User.Read`, `Files.ReadWrite`, `Notes.Create` and `Notes.Read`. The redirect URI is derived from the current origin and Vite base URL.

The frontend obtains Graph tokens and uses Graph directly for browser-side OneDrive and OneNote operations. It passes a delegated Graph token to the backend for backend-side OneDrive downloads and uploads.

## Backend

The Python FastAPI application is in `backend/main.py`. The inspected source reads:

- `ASSEMBLYAI_API_KEY`
- `GEMINI_API_KEY`
- `GEMINI_MODEL` (source default: `gemini-3.5-flash-lite`)
- `GEMINI_TIMEOUT_MS`
- `FFMPEG_PATH`
- `MAX_DOWNLOAD_BYTES` (source default: 3 GiB)
- `CORS_ORIGINS`

The inspected default CORS origins include localhost and 127.0.0.1 on ports 5173 and 5174. `CORS_ORIGINS` can override them.

## API endpoints identified

| Endpoint | Purpose |
|---|---|
| `GET /health` | Health and configuration status |
| `GET /diagnostics/assemblyai` | Read-only AssemblyAI credential test |
| `GET /diagnostics/pdf` | Check PDF library availability |
| `POST /test-gemini` | Gemini connectivity test |
| `POST /transcribe-onedrive` | Download recording and transcribe |
| `POST /summarise-and-save` | Save transcript, generate and save summary |
| `POST /generate-missing-summary` | Generate summary from existing transcript |
| `POST /render-summary-pdf` | Render Markdown summary to PDF and save it |

Check `backend/main.py` for authoritative request/response models before changing an API.

## Processing

### Transcription

The frontend sends OneDrive item ID, filename and Graph token to `/transcribe-onedrive`. The backend downloads the file. MP4 is sent directly to AssemblyAI. For MKV, FFmpeg extracts mono 16 kHz MP3 audio at 64 kbps. The backend returns transcript text.

### Summary

The frontend sends transcript text, filename, Graph token and destination folder ID to `/summarise-and-save`. The backend saves the transcript first, calls Gemini to generate student-facing Markdown revision notes, then saves the summary. The Gemini prompt specifies sections for overview, key concepts, worked examples, key points and homework/next steps, with instructions to preserve mathematical notation and not invent unsupported material.

### PDF and OneNote

The frontend sends Markdown to `/render-summary-pdf`. The backend saves the generated PDF to OneDrive. The frontend creates anonymous view links for the recording and PDF, downloads the PDF and submits multipart HTML plus PDF to the Microsoft Graph OneNote pages endpoint. The HTML references the PDF as a renderable data part, allowing OneNote to display its pages as images.

## Security boundaries

- Microsoft Graph access is delegated through the signed-in user.
- API keys belong in backend environment configuration, never frontend source.
- `.env` and virtual environments are ignored by the inspected `.gitignore`.
- Anonymous sharing links are created for the recording and PDF. Review this before production use.
- Restrict production CORS to intended frontend origins.
