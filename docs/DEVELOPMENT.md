# Development

## Prerequisites

The inspected development environment is Windows 11. The project requires Git, Node.js/npm, Python (the backend startup script prefers Python 3.13), and FFmpeg for MKV audio extraction. You also need access to the Microsoft, AssemblyAI and Google services.

## Local repository

The working copy used during development:

`C:\Users\clive\Dev\Tutor-automate-OneNote`

The branch observed during documentation work was `development`.

## Frontend

From the repository root:

```powershell
npm install
npm run dev
```

Vite prints the local development URL. The backend's inspected CORS defaults include ports 5173 and 5174.

Check `package.json` for the available scripts before using build or lint commands.

## Backend

Run from the repository root:

```powershell
.\backend\Run-Backend.bat
```

The batch file changes into `backend`, creates `.venv` if needed (preferring Python 3.13), activates it, upgrades pip, installs `requirements.txt`, creates `.env` from `.env.example` if absent, and starts Uvicorn at `http://127.0.0.1:8000`.

Health check:

`http://127.0.0.1:8000/health`

The current inspected `main.py` uses AssemblyAI and Gemini. Older Ollama references remain in some legacy backend setup text/configuration and should be reviewed.

## Environment configuration

`backend/.env.example` is the template. The current source reads `ASSEMBLYAI_API_KEY`, `GEMINI_API_KEY`, `GEMINI_MODEL`, `GEMINI_TIMEOUT_MS`, `FFMPEG_PATH`, `MAX_DOWNLOAD_BYTES` and `CORS_ORIGINS`.

Do not put real keys in source code, documentation, GitHub issues or chat messages. The inspected `.gitignore` excludes `.env` and `.env.*` except `.env.example`, as well as Python virtual environments and bytecode.

## Suggested testing

1. Start the backend and check `/health`.
2. Use `/diagnostics/assemblyai` if transcription credentials are suspect.
3. Use `/test-gemini` if summary generation connectivity is suspect.
4. Use `/diagnostics/pdf` if PDF generation fails.
5. Run the frontend locally and sign in with a test Microsoft account.
6. Test using a short, non-confidential recording in a test OneDrive folder.
7. Check transcript, summary and PDF uploads.
8. Create a OneNote page in a test notebook and section.

Never include API keys, bearer tokens or private lesson content in diagnostic reports.

## VSCodium Markdown preview

- Open a `.md` file.
- `Ctrl+Shift+V` opens Markdown preview.
- `Ctrl+K`, then `V` opens preview to the side.
- Save and inspect the rendered preview before committing.
