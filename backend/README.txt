STAGE 3 BACKEND

This is the existing Stage 2 FastAPI backend with one additional endpoint:
POST /summarise-and-save

It sends the transcript to local Ollama (qwen3:4b), then uploads a transcript
TXT file and a student revision summary TXT file to the same OneDrive folder as
the source recording.

Keep the existing .env file containing ASSEMBLYAI_API_KEY. No additional API key
is needed. Ollama must be running locally at http://127.0.0.1:11434 and the
qwen3:4b model must be installed.

Start with Run-Backend.bat. Health check: http://127.0.0.1:8000/health
