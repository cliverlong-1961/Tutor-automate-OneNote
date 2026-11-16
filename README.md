# Tutor Video OneNote

A React/TypeScript web application for turning tutoring recordings into transcripts, student revision notes, PDFs and OneNote lesson pages.

## Main workflow

1. Sign in with Microsoft.
2. Browse to a OneDrive folder and select an MP4 or MKV recording.
3. The FastAPI backend downloads the recording from OneDrive and sends it to AssemblyAI for transcription.
4. The transcript is returned to the frontend.
5. The frontend sends the transcript to the backend; Google Gemini generates student-facing Markdown revision notes.
6. The backend saves the transcript and summary as text files in OneDrive.
7. The backend renders the summary as a PDF and saves it to OneDrive.
8. The frontend creates view-sharing links for the recording and PDF.
9. The user selects a OneNote notebook and section.
10. The app creates a OneNote page with links to the recording and PDF and embeds the PDF pages as images.

The transcript is retained in OneDrive but is not linked on the OneNote page.

## Technology

- Frontend: React, TypeScript, Vite
- Authentication and Microsoft services: MSAL and Microsoft Graph
- Backend: Python, FastAPI and Uvicorn
- Transcription: AssemblyAI
- Summary generation: Google Gemini
- Hosting: GitHub Pages (frontend) and Azure Container Apps (backend)

## Documentation

- [User guide](docs/USER-GUIDE.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Development](docs/DEVELOPMENT.md)
- [Deployment](docs/DEPLOYMENT.md)
- [Troubleshooting](docs/TROUBLESHOOTING.md)
- [Changelog](docs/CHANGELOG.md)
- [Open questions](docs/OPEN-QUESTIONS.md)

## Security

Do not commit `backend/.env`, API keys or access tokens. The inspected `.gitignore` excludes `.env` files (except `.env.example`) and Python virtual environments.

**Privacy warning:** the frontend requests anonymous view links for the recording and PDF (`scope: anonymous`). Anyone who obtains those links may be able to view the files. Review this setting before using the app with confidential student recordings.

This is an initial source-informed documentation set. Items not yet verified against the complete deployment configuration are listed in [Open questions](docs/OPEN-QUESTIONS.md).
