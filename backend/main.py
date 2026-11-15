import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

import assemblyai as aai
import requests
from google import genai
from google.genai import types
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from io import BytesIO
from xml.sax.saxutils import escape as xml_escape

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

API_KEY = os.getenv("ASSEMBLYAI_API_KEY", "").strip()
FFMPEG_PATH = os.getenv("FFMPEG_PATH", "ffmpeg")
MAX_DOWNLOAD_BYTES = int(os.getenv("MAX_DOWNLOAD_BYTES", str(3 * 1024 * 1024 * 1024)))
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()
GEMINI_TIMEOUT_MS = int(os.getenv("GEMINI_TIMEOUT_MS", "180000"))

app = FastAPI(title="Tutoring Automation Transcription API")

# Comma-separated browser origins, e.g. https://username.github.io.
# Local development origins remain enabled by default.
DEFAULT_CORS_ORIGINS = (
    "http://localhost:5174,"
    "http://127.0.0.1:5174,"
    "http://localhost:5173,"
    "http://127.0.0.1:5173"
)
CORS_ORIGINS = [
    origin.strip()
    for origin in os.getenv("CORS_ORIGINS", DEFAULT_CORS_ORIGINS).split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=False,
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)


class TranscribeRequest(BaseModel):
    item_id: str = Field(min_length=1, max_length=500)
    file_name: str = Field(min_length=1, max_length=500)
    graph_access_token: str = Field(min_length=20)


@app.get("/health")
def health():
    return {
        "status": "ok",
        "assemblyai_key_configured": bool(API_KEY),
        "gemini_api_key_configured": bool(GEMINI_API_KEY),
        "gemini_model": GEMINI_MODEL,
    }


@app.get("/diagnostics/assemblyai")
def diagnostics_assemblyai():
    """Validate AssemblyAI credentials with a read-only account API request."""
    if not API_KEY:
        raise HTTPException(status_code=503, detail="ASSEMBLYAI_API_KEY is not configured.")
    try:
        response = requests.get(
            "https://api.assemblyai.com/v2/transcript",
            headers={"authorization": API_KEY},
            params={"limit": 1},
            timeout=15,
        )
        if response.status_code == 200:
            return {
                "ok": True,
                "service": "AssemblyAI",
                "message": "API authentication and read-only transcript-list request succeeded.",
            }
        if response.status_code in (401, 403):
            raise HTTPException(
                status_code=502,
                detail={"message": "AssemblyAI rejected the configured API key.", "upstream_status": response.status_code},
            )
        raise HTTPException(
            status_code=502,
            detail={"message": "AssemblyAI returned an unexpected response to the diagnostic request.", "upstream_status": response.status_code},
        )
    except requests.Timeout as exc:
        raise HTTPException(status_code=504, detail="AssemblyAI did not respond within 15 seconds.") from exc
    except requests.RequestException as exc:
        raise HTTPException(status_code=502, detail=f"Could not reach AssemblyAI: {type(exc).__name__}.") from exc


@app.get("/diagnostics/pdf")
def diagnostics_pdf():
    """Check that the server-side PDF rendering dependencies are importable."""
    import importlib.util

    missing = [name for name in ("matplotlib", "reportlab") if importlib.util.find_spec(name) is None]
    if missing:
        raise HTTPException(status_code=503, detail=f"PDF dependencies missing: {', '.join(missing)}.")
    return {
        "ok": True,
        "service": "PDF generation",
        "message": "Matplotlib and ReportLab are installed and importable.",
    }


@app.post("/test-gemini")
def test_gemini_connectivity():
    """Make a minimal live request to Gemini and return useful diagnostic details."""
    if not GEMINI_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="GEMINI_API_KEY is not configured in the Azure Container App environment.",
        )

    try:
        client = genai.Client(
            api_key=GEMINI_API_KEY,
            http_options=types.HttpOptions(timeout=GEMINI_TIMEOUT_MS),
        )
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents="Reply with exactly: GEMINI_TEST_OK",
            config=types.GenerateContentConfig(temperature=0, max_output_tokens=20),
        )
        response_text = (response.text or "").strip()
        if "GEMINI_TEST_OK" not in response_text:
            raise RuntimeError(
                "Gemini responded, but its response did not contain the expected text GEMINI_TEST_OK. "
                f"Response received: {response_text[:300]!r}"
            )
        return {
            "ok": True,
            "model": GEMINI_MODEL,
            "message": "Gemini returned the expected test response.",
        }
    except Exception as exc:
        error_type = type(exc).__name__
        # google-genai/httpx exceptions may expose an upstream HTTP status on
        # code, status_code, or response.status_code. Report it when available.
        upstream_status = getattr(exc, "status_code", None)
        if upstream_status is None:
            upstream_status = getattr(exc, "code", None)
        response_obj = getattr(exc, "response", None)
        if upstream_status is None and response_obj is not None:
            upstream_status = getattr(response_obj, "status_code", None)

        message = str(exc).strip() or repr(exc)
        message = message[:1800]
        is_timeout = (
            "timeout" in error_type.lower()
            or "timeout" in message.lower()
            or isinstance(exc, TimeoutError)
        )
        if is_timeout:
            detail = (
                f"No response from Gemini before the configured timeout of {GEMINI_TIMEOUT_MS} ms. "
                f"Exception: {error_type}. {message}"
            )
            raise HTTPException(
                status_code=504,
                detail={
                    "message": detail,
                    "error_type": error_type,
                    "gemini_status_code": upstream_status,
                    "model": GEMINI_MODEL,
                    "timeout_ms": GEMINI_TIMEOUT_MS,
                },
            ) from exc

        detail = {
            "message": f"Gemini request failed. {error_type}: {message}",
            "error_type": error_type,
            "gemini_status_code": upstream_status,
            "model": GEMINI_MODEL,
        }
        raise HTTPException(status_code=502, detail=detail) from exc


def safe_filename(name: str) -> str:
    # Use only the basename; never allow a client-supplied path.
    name = Path(name.replace("\\", "/")).name
    name = re.sub(r"[^A-Za-z0-9._ ()-]", "_", name).strip(" .")
    return name or "recording.mp4"


def _graph_error_detail(response: requests.Response, action: str) -> dict:
    """Return useful Graph diagnostics without exposing HTML error pages or tokens."""
    try:
        payload = response.json()
        graph_error = payload.get("error", {}) if isinstance(payload, dict) else {}
        code = graph_error.get("code") if isinstance(graph_error, dict) else None
        message = graph_error.get("message") if isinstance(graph_error, dict) else None
    except (ValueError, requests.exceptions.JSONDecodeError):
        code = None
        message = None
    request_id = response.headers.get("request-id") or response.headers.get("client-request-id")
    date = response.headers.get("Date")
    if not message:
        message = "Microsoft returned a non-JSON response. The response body was omitted because it may be an HTML sign-in or service error page."
    return {
        "message": f"Microsoft Graph could not {action} (HTTP {response.status_code}). {message}",
        "graph_status_code": response.status_code,
        "graph_error_code": code,
        "request_id": request_id,
        "response_date": date,
    }


def download_from_onedrive(item_id: str, token: str, destination: Path) -> None:
    """Validate the Graph token/item, then download using Graph's temporary URL."""
    encoded_id = requests.utils.quote(item_id, safe="")
    base_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{encoded_id}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    try:
        # Check that the delegated token is valid and that this item is accessible.
        metadata_response = requests.get(
            base_url,
            headers=headers,
            params={"$select": "id,name,size,file,parentReference"},
            timeout=(20, 45),
        )
        if not metadata_response.ok:
            status = 401 if metadata_response.status_code == 401 else 403 if metadata_response.status_code == 403 else 502
            raise HTTPException(
                status_code=status,
                detail=_graph_error_detail(metadata_response, "verify access to the selected OneDrive recording"),
            )

        # /content normally returns a redirect to a short-lived download URL.
        # Handle that redirect explicitly so the Graph bearer token is never
        # forwarded to the OneDrive/SharePoint download host.
        content_response = requests.get(
            f"{base_url}/content",
            headers=headers,
            stream=True,
            allow_redirects=False,
            timeout=(20, 60),
        )
        if content_response.status_code in (301, 302, 303, 307, 308):
            download_url = content_response.headers.get("Location")
            content_response.close()
            if not download_url:
                raise HTTPException(
                    status_code=502,
                    detail="Microsoft Graph redirected the recording download but did not provide a download URL.",
                )
            # The URL is pre-authorised by Microsoft; do not attach the Graph token.
            response = requests.get(
                download_url,
                stream=True,
                headers={"User-Agent": "TutoringAutomation/1.0"},
                timeout=(20, 180),
            )
        else:
            response = content_response

        with response:
            if response.status_code != 200:
                status = 401 if response.status_code == 401 else 403 if response.status_code == 403 else 502
                action = "download the recording from OneDrive" if response.url != f"{base_url}/content" else "retrieve the recording content"
                raise HTTPException(
                    status_code=status,
                    detail=_graph_error_detail(response, action),
                )

            total = 0
            with destination.open("wb") as output:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if not chunk:
                        continue
                    total += len(chunk)
                    if total > MAX_DOWNLOAD_BYTES:
                        raise HTTPException(
                            status_code=413,
                            detail=f"Recording exceeds the configured {MAX_DOWNLOAD_BYTES // (1024 * 1024)} MB download limit.",
                        )
                    output.write(chunk)
    except HTTPException:
        raise
    except requests.Timeout as exc:
        raise HTTPException(
            status_code=504,
            detail=f"Timed out while downloading the recording from OneDrive: {type(exc).__name__}: {exc}",
        ) from exc
    except requests.RequestException as exc:
        raise HTTPException(
            status_code=502,
            detail=f"OneDrive download request failed: {type(exc).__name__}: {exc}",
        ) from exc


@app.post("/transcribe-onedrive")
def transcribe_onedrive(request: TranscribeRequest):
    if not API_KEY:
        raise HTTPException(
            status_code=500,
            detail="AssemblyAI API key is not configured. Add ASSEMBLYAI_API_KEY to backend/.env and restart the backend.",
        )

    filename = safe_filename(request.file_name)
    extension = Path(filename).suffix.lower()
    if extension not in {".mp4", ".mkv"}:
        raise HTTPException(status_code=400, detail="Only MP4 and MKV recordings are supported.")

    with tempfile.TemporaryDirectory(prefix="tutoring-transcription-") as temp_dir:
        work_dir = Path(temp_dir)
        original_path = work_dir / filename
        download_from_onedrive(request.item_id, request.graph_access_token, original_path)

        transcription_path = original_path
        if extension == ".mkv":
            if not shutil.which(FFMPEG_PATH) and not Path(FFMPEG_PATH).exists():
                raise HTTPException(
                    status_code=500,
                    detail="FFmpeg was not found. Install FFmpeg and ensure ffmpeg.exe is on PATH, then restart the backend.",
                )
            audio_path = work_dir / f"{Path(filename).stem}.mp3"
            try:
                subprocess.run(
                    [
                        FFMPEG_PATH, "-y", "-i", str(original_path),
                        "-vn", "-ac", "1", "-ar", "16000", "-b:a", "64k",
                        str(audio_path),
                    ],
                    check=True,
                    capture_output=True,
                    text=True,
                    timeout=60 * 60,
                )
            except subprocess.TimeoutExpired as exc:
                raise HTTPException(status_code=504, detail="FFmpeg audio extraction timed out.") from exc
            except subprocess.CalledProcessError as exc:
                detail = (exc.stderr or str(exc))[-2000:]
                raise HTTPException(status_code=422, detail=f"Could not extract audio from MKV: {detail}") from exc
            transcription_path = audio_path

        try:
            aai.settings.api_key = API_KEY
            transcriber = aai.Transcriber()
            transcript = transcriber.transcribe(str(transcription_path))
            if transcript.status == aai.TranscriptStatus.error:
                raise HTTPException(
                    status_code=502,
                    detail=f"AssemblyAI transcription failed: {transcript.error or 'Unknown error'}",
                )
            if not transcript.text:
                raise HTTPException(status_code=502, detail="AssemblyAI completed but returned an empty transcript.")
            return {
                "transcript_id": transcript.id,
                "status": str(transcript.status),
                "file_name": filename,
                "text": transcript.text,
            }
        except HTTPException:
            raise
        except Exception as exc:
            raise HTTPException(status_code=502, detail=f"AssemblyAI request failed: {exc}") from exc


class SaveLessonNotesRequest(BaseModel):
    file_name: str = Field(min_length=1, max_length=500)
    transcript_text: str = Field(min_length=1, max_length=2_000_000)
    graph_access_token: str = Field(min_length=20)
    folder_id: Optional[str] = Field(default=None, max_length=500)


def upload_text_to_onedrive(filename: str, content: str, token: str, folder_id: Optional[str]):
    encoded_name = requests.utils.quote(filename, safe="")
    if folder_id:
        encoded_folder = requests.utils.quote(folder_id, safe="")
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{encoded_folder}:/{encoded_name}:/content"
    else:
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_name}:/content"
    try:
        response = requests.put(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "text/plain; charset=utf-8",
            },
            data=content.encode("utf-8"),
            timeout=(20, 180),
        )
    except requests.RequestException as exc:
        raise HTTPException(status_code=502, detail=f"OneDrive upload failed for {filename}: {exc}") from exc
    if response.status_code not in (200, 201):
        raise HTTPException(
            status_code=502,
            detail=f"Microsoft Graph could not save {filename} ({response.status_code}): {response.text[:1000]}",
        )
    return response.json()


def generate_student_summary(transcript_text: str, recording_name: str) -> str:
    if not GEMINI_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="Gemini API key is not configured. Add GEMINI_API_KEY to backend/.env and restart the backend.",
        )

    prompt = f"""You are preparing clear, accurate revision notes for a student from a tutoring lesson transcript.

Create a student-facing lesson summary in Markdown. Use only information supported by the transcript. Do not invent facts, examples, answers, homework, or topics. Preserve the teacher's explanations and the examples actually discussed. If homework or follow-up work was not explicitly set, write 'No specific homework was identified in the transcript.' If a worked example is unclear or incomplete, say so rather than reconstructing missing steps.

Use this structure. Do not add a document-level title or a first-line heading such as "# Lesson revision notes"; begin directly with the first section heading:
## Lesson overview
A short description of what the lesson covered.
## Key concepts
Concise explanations of the main ideas, definitions, rules and formulae.
## Worked examples
Summarise examples actually worked through, including important steps and conclusions. If none are identifiable, say so.
## Key points to remember
A short bullet list of the most useful takeaways.
## Homework and next steps
List homework or follow-up tasks explicitly mentioned. Do not infer homework.

Mathematical formatting:
- Use LaTeX for mathematical expressions, not spoken descriptions such as 'dy by dx'.
- Use $...$ for inline mathematics. Every opening $ must have exactly one matching closing $ on the same line.
- Use $$...$$ for displayed equations, with matching opening and closing $$ delimiters. Put displayed equations on separate lines.
- Never mix single-dollar and double-dollar delimiters, and never leave a dollar delimiter unmatched.
- Do not put ordinary prose, Markdown bullets, headings or explanations inside a LaTeX math expression.
- Check the complete summary for balanced and correctly paired math delimiters before returning it.
- Preserve signs, indices, powers, brackets, fractions and variable names accurately.
- Do not create or guess mathematical expressions that are not supported by the transcript.

Write for a student revising independently: clear language, concise explanations, and enough detail to be useful. Remove conversational filler, greetings, apologies and technical meeting chatter. Do not mention that this was generated by AI.

Recording: {recording_name}

TRANSCRIPT:
{transcript_text}
"""

    try:
        client = genai.Client(
            api_key=GEMINI_API_KEY,
            http_options=types.HttpOptions(timeout=GEMINI_TIMEOUT_MS),
        )
    except Exception as exc:
        raise HTTPException(
            status_code=502,
            detail=(
                f"Gemini connectivity test failed for {GEMINI_MODEL}: "
                f"could not initialise the Gemini client ({type(exc).__name__}: {exc}). "
                "The lesson summary was not attempted."
            ),
        ) from exc

    # First make a very small request. This checks the configured key, model
    # and Gemini service before sending the much larger lesson transcript.
    try:
        test_response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents="Reply with exactly: GEMINI_TEST_OK",
            config=types.GenerateContentConfig(
                temperature=0,
                max_output_tokens=20,
            ),
        )
        test_text = (test_response.text or "").strip()
        if "GEMINI_TEST_OK" not in test_text:
            raise RuntimeError(
                "Gemini responded, but did not return the expected connectivity-test text."
            )
    except Exception as exc:
        # Stop here: do not send the full transcript if the preliminary test fails.
        raise HTTPException(
            status_code=502,
            detail=(
                f"Gemini connectivity test failed using {GEMINI_MODEL}: "
                f"{type(exc).__name__}: {exc}. "
                "The lesson summary was not attempted. Please try again later."
            ),
        ) from exc

    # The short test succeeded. Now submit the full transcript for summarisation.
    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                temperature=0.2,
                max_output_tokens=8192,
            ),
        )
        summary = (response.text or "").strip()
    except Exception as exc:
        # Keep the API key and request headers out of the returned error.
        raise HTTPException(
            status_code=502,
            detail=f"Gemini summarisation failed using {GEMINI_MODEL} after a successful connectivity test: {type(exc).__name__}: {exc}",
        ) from exc

    if not summary:
        raise HTTPException(status_code=502, detail="Gemini returned an empty lesson summary.")

    # Defensively remove the old document-level title if the model adds it anyway.
    summary_lines = summary.splitlines()
    while summary_lines and not summary_lines[0].strip():
        summary_lines.pop(0)
    if summary_lines and summary_lines[0].strip().lower() in {
        "# lesson revision notes", "lesson revision notes", "# lesson summary"
    }:
        summary_lines.pop(0)
        while summary_lines and not summary_lines[0].strip():
            summary_lines.pop(0)
    return "\n".join(summary_lines).strip()


@app.post("/summarise-and-save")
def summarise_and_save(request: SaveLessonNotesRequest):
    filename = safe_filename(request.file_name)
    stem = Path(filename).stem
    transcript_filename = f"{stem}_transcript.txt"
    summary_filename = f"{stem}_summary.txt"

    # Save the transcript first. It remains available in OneDrive even if
    # the Gemini request fails or times out.
    transcript_item = upload_text_to_onedrive(
        transcript_filename, request.transcript_text, request.graph_access_token, request.folder_id
    )

    summary = generate_student_summary(request.transcript_text, filename)
    summary_item = upload_text_to_onedrive(
        summary_filename, summary, request.graph_access_token, request.folder_id
    )
    return {
        "transcript": request.transcript_text,
        "summary": summary,
        "transcript_file": transcript_item,
        "summary_file": summary_item,
        "message": f"Saved {transcript_filename} and {summary_filename} to OneDrive.",
    }


class GenerateMissingSummaryRequest(BaseModel):
    file_name: str = Field(min_length=1, max_length=500)
    transcript_text: str = Field(min_length=1, max_length=2_000_000)
    graph_access_token: str = Field(min_length=20)
    folder_id: Optional[str] = Field(default=None, max_length=500)


@app.post("/generate-missing-summary")
def generate_missing_summary(request: GenerateMissingSummaryRequest):
    """Generate and save only the summary, using an existing OneDrive transcript."""
    filename = safe_filename(request.file_name)
    stem = Path(filename).stem
    summary_filename = f"{stem}_summary.txt"

    summary = generate_student_summary(request.transcript_text, filename)
    summary_item = upload_text_to_onedrive(
        summary_filename, summary, request.graph_access_token, request.folder_id
    )
    return {
        "summary": summary,
        "summary_file": summary_item,
        "message": f"Generated and saved {summary_filename} to OneDrive using the existing transcript. No video transcription was performed.",
    }


class RenderSummaryPdfRequest(BaseModel):
    file_name: str = Field(min_length=1, max_length=500)
    markdown_text: str = Field(min_length=1, max_length=2_000_000)
    graph_access_token: str = Field(min_length=20)
    folder_id: Optional[str] = Field(default=None, max_length=500)


def upload_pdf_to_onedrive(filename: str, pdf_bytes: bytes, token: str, folder_id: Optional[str]):
    encoded_name = requests.utils.quote(filename, safe="")
    if folder_id:
        encoded_folder = requests.utils.quote(folder_id, safe="")
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{encoded_folder}:/{encoded_name}:/content"
    else:
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{encoded_name}:/content"
    try:
        response = requests.put(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/pdf",
            },
            data=pdf_bytes,
            timeout=(20, 240),
        )
    except requests.RequestException as exc:
        raise HTTPException(status_code=502, detail=f"OneDrive PDF upload failed: {exc}") from exc
    if response.status_code not in (200, 201):
        raise HTTPException(
            status_code=502,
            detail=f"Microsoft Graph could not save {filename} ({response.status_code}): {response.text[:1000]}",
        )
    return response.json()


def normalise_latex_for_mathtext(latex: str) -> str:
    """Convert common full-LaTeX commands and shorthand into MathText syntax.

    This deliberately normalises known equivalents rather than deleting
    unknown commands, so mathematical meaning is not silently lost.
    """
    expr = latex.strip()

    # Commands used frequently by Gemini and conventional LaTeX documents
    # that have equivalent names in Matplotlib MathText.
    replacements = {
        r"\implies": r"\Rightarrow",
        r"\Longrightarrow": r"\Rightarrow",
        r"\longrightarrow": r"\rightarrow",
        r"\to": r"\rightarrow",
        r"\iff": r"\Leftrightarrow",
        r"\Longleftrightarrow": r"\Leftrightarrow",
        r"\longleftrightarrow": r"\leftrightarrow",
        r"\dfrac": r"\frac",
        r"\tfrac": r"\frac",
        r"\geqslant": r"\geq",
        r"\leqslant": r"\leq",
        r"\land": r"\wedge",
        r"\lor": r"\vee",
        r"\lnot": r"\neg",
    }
    # Replace commands as complete control words (not as prefixes).
    for source, target in replacements.items():
        expr = re.sub(re.escape(source) + r"(?![A-Za-z])", lambda _m, t=target: t, expr)

    # Matplotlib MathText uses \\leq and \\geq, not the shorter LaTeX forms.
    expr = re.sub(r"\\le(?![a-zA-Z])", r"\\leq", expr)
    expr = re.sub(r"\\ge(?![a-zA-Z])", r"\\geq", expr)

    # MathText does not need LaTeX's sizing delimiters.
    expr = re.sub(r"\\(?:displaystyle|textstyle|scriptstyle|scriptscriptstyle)\b", "", expr)
    expr = re.sub(r"\\(?:left|right)\b", "", expr)

    # Matplotlib accepts \mathrm{...}; use it for operator names.
    expr = re.sub(r"\\operatorname\s*\{([^{}]*)\}", r"\\mathrm{\1}", expr)

    # Gemini sometimes emits shorthand such as \frac12 or \sqrt3.
    # Convert simple single-character arguments to explicit braced arguments.
    expr = re.sub(r"\\frac\s*([A-Za-z0-9])\s*([A-Za-z0-9])",
                  r"\\frac{\1}{\2}", expr)
    expr = re.sub(r"\\sqrt\s*([A-Za-z0-9])", r"\\sqrt{\1}", expr)

    # MathText expects a braced argument for these common commands.
    expr = re.sub(r"\\(?:mathrm|mathbf|mathit|mathcal|mathbb)\s+([A-Za-z])",
                  lambda m: "\\" + m.group(0).split()[0][1:] + "{" + m.group(1) + "}", expr)
    return expr



def validate_markdown_math_delimiters(markdown_text: str) -> None:
    """Reject unmatched or mixed $ / $$ math delimiters before PDF rendering.

    Fenced code blocks are ignored. Inline $...$ expressions must close on
    the same line; display $$...$$ expressions may span multiple lines.
    """
    opening = None
    in_fenced_code = False

    for line_number, line in enumerate(markdown_text.splitlines(), start=1):
        if re.match(r"^\s*(```|~~~)", line):
            in_fenced_code = not in_fenced_code
            continue
        if in_fenced_code:
            continue

        i = 0
        while i < len(line):
            if line[i] != "$":
                i += 1
                continue
            backslashes = 0
            j = i - 1
            while j >= 0 and line[j] == "\\":
                backslashes += 1
                j -= 1
            if backslashes % 2:
                i += 1
                continue

            delimiter = "$$" if line.startswith("$$", i) else "$"
            if opening is None:
                opening = (delimiter, line_number, i + 1)
            elif opening[0] != delimiter:
                opened, opened_line, opened_col = opening
                raise HTTPException(
                    status_code=422,
                    detail=(
                        f"Invalid LaTeX delimiters: found {delimiter} at line "
                        f"{line_number}, column {i + 1}, but {opened} was opened "
                        f"at line {opened_line}, column {opened_col}. Use matching "
                        f"$...$ for inline maths or $$...$$ for display equations."
                    ),
                )
            else:
                opening = None
            i += len(delimiter)

            if opening is not None and opening[0] == "$":
                if not re.search(r"(?<!\\)\$", line[i:]):
                    _, opened_line, opened_col = opening
                    raise HTTPException(
                        status_code=422,
                        detail=f"Unclosed inline LaTeX delimiter at line {opened_line}, column {opened_col}. Use a matching $...$ pair on the same line.",
                    )

    if opening is not None:
        delimiter, opened_line, opened_col = opening
        raise HTTPException(
            status_code=422,
            detail=f"Unclosed {delimiter} LaTeX delimiter at line {opened_line}, column {opened_col}. Check matching math delimiters.",
        )


def build_revision_pdf(markdown_text: str, lesson_title: str) -> bytes:
    """Create an A4 revision-notes PDF; LaTeX is rendered to crisp PNG equation assets."""
    try:
        import mistune
        import matplotlib
        matplotlib.use("Agg")
        from matplotlib import mathtext
        from matplotlib.font_manager import FontProperties, findfont
        from bs4 import BeautifulSoup, NavigableString
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, PageBreak,
            ListFlowable, ListItem, Table, TableStyle, Preformatted, HRFlowable,
        )
        from reportlab.lib.utils import ImageReader
    except ImportError as exc:
        raise HTTPException(
            status_code=500,
            detail="PDF dependencies are missing. In backend\\.venv run: python -m pip install reportlab matplotlib mistune beautifulsoup4",
        ) from exc

    # Validate math delimiters before Markdown parsing or Matplotlib rendering.
    validate_markdown_math_delimiters(markdown_text)

    safe_stem = Path(safe_filename(lesson_title)).stem
    with tempfile.TemporaryDirectory(prefix="tutoring-pdf-") as temp_dir:
        work = Path(temp_dir)
        equation_counter = 0
        equation_files = {}

        def render_equation(latex: str, display: bool) -> Path:
            nonlocal equation_counter
            equation_counter += 1
            image_path = work / f"equation_{equation_counter}.png"
            expr = normalise_latex_for_mathtext(latex)
            try:
                mathtext.math_to_image(
                    f"${expr}$",
                    str(image_path),
                    prop=FontProperties(size=12),  # 2 pt larger than the 10 pt explanatory text
                    dpi=240,
                    format="png",
                    color="#172554",
                )
            except Exception as exc:
                raise HTTPException(
                    status_code=422,
                    detail=f"Could not render a LaTeX expression in the lesson summary: {expr[:180]}. Simplify this expression or use standard LaTeX commands supported by Matplotlib mathtext. Details: {type(exc).__name__}: {exc}",
                ) from exc
            equation_files[str(equation_counter)] = image_path
            return image_path

        # Extract display equations first so Markdown treats them as separate blocks.
        display_tokens = {}
        def hold_display(match):
            key = f"REVISIONDISPLAYMATH{len(display_tokens)}TOKEN"
            display_tokens[key] = match.group(1).strip()
            return f"\n\n{key}\n\n"
        prepared = re.sub(r"\$\$(.+?)\$\$", hold_display, markdown_text, flags=re.S)
        prepared = re.sub(r"\\\[(.+?)\\\]", hold_display, prepared, flags=re.S)

        # Protect inline maths before converting Markdown.
        inline_tokens = {}
        def hold_inline(match):
            key = f"REVISIONINLINEMATH{len(inline_tokens)}TOKEN"
            inline_tokens[key] = (match.group(1) or match.group(2)).strip()
            return key
        prepared = re.sub(r"\\\((.+?)\\\)|(?<!\$)\$([^$\n]+)\$(?!\$)", hold_inline, prepared)

        markdown_renderer = mistune.create_markdown(plugins=["table", "strikethrough"])
        html = markdown_renderer(prepared)
        soup = BeautifulSoup(html, "html.parser")

        # Replace inline math placeholders in text nodes with real inline PNG images.
        for text_node in list(soup.find_all(string=True)):
            if not isinstance(text_node, NavigableString) or not any(k in str(text_node) for k in inline_tokens):
                continue
            original = str(text_node)
            fragments = re.split("(" + "|".join(map(re.escape, inline_tokens.keys())) + ")", original)
            for fragment in fragments:
                if not fragment:
                    continue
                if fragment in inline_tokens:
                    img_path = render_equation(inline_tokens[fragment], False)
                    reader = ImageReader(str(img_path))
                    width_px, height_px = reader.getSize()
                    # Rendered at 12 pt against 10 pt explanatory text.
                    # Convert the 240-dpi PNG dimensions to PDF points without
                    # applying the former 13 pt height cap: that cap scaled many
                    # inline expressions down and made their apparent size vary.
                    # Keep the natural 12 pt math size and current inline position.
                    natural_scale = 72 / 240
                    width_pt = width_px * natural_scale
                    height_pt = height_px * natural_scale
                    # Only very long inline formulae are scaled down enough to
                    # fit the text measure; ordinary formulae remain at 12 pt.
                    if width_pt > 260:
                        scale = 260 / width_pt
                        width_pt *= scale
                        height_pt *= scale
                    img = soup.new_tag("img", src=str(img_path), width=f"{width_pt:.1f}", height=f"{height_pt:.1f}", valign="-2")
                    text_node.insert_before(img)
                else:
                    text_node.insert_before(NavigableString(fragment))
            text_node.extract()

        # ReportLab fonts: use DejaVu for Greek symbols and broad Unicode coverage.
        font_regular = findfont(FontProperties(family="DejaVu Sans", weight="normal"))
        font_bold = findfont(FontProperties(family="DejaVu Sans", weight="bold"))
        font_italic = findfont(FontProperties(family="DejaVu Sans", style="oblique"))
        font_bold_italic = findfont(FontProperties(family="DejaVu Sans", weight="bold", style="oblique"))
        for name, path in [("DejaVu", font_regular), ("DejaVu-Bold", font_bold),
                           ("DejaVu-Italic", font_italic), ("DejaVu-BoldItalic", font_bold_italic)]:
            if name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(name, path))
        pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold",
                                      italic="DejaVu-Italic", boldItalic="DejaVu-BoldItalic")

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name="RevTitle", fontName="DejaVu-Bold", fontSize=21, leading=27,
                                  textColor=colors.HexColor("#17365D"), spaceAfter=12))
        styles.add(ParagraphStyle(name="RevH1", fontName="DejaVu-Bold", fontSize=16, leading=20,
                                  textColor=colors.HexColor("#17365D"), spaceBefore=10, spaceAfter=8))
        styles.add(ParagraphStyle(name="RevH2", fontName="DejaVu-Bold", fontSize=13, leading=17,
                                  textColor=colors.HexColor("#245A8D"), spaceBefore=9, spaceAfter=6))
        styles.add(ParagraphStyle(name="RevH3", fontName="DejaVu-Bold", fontSize=11, leading=14,
                                  textColor=colors.HexColor("#334155"), spaceBefore=7, spaceAfter=4))
        styles.add(ParagraphStyle(name="RevBody", fontName="DejaVu", fontSize=10, leading=14.5,
                                  textColor=colors.HexColor("#202938"), spaceAfter=6))
        styles.add(ParagraphStyle(name="RevBullet", parent=styles["RevBody"], leftIndent=13,
                                  firstLineIndent=-8, bulletIndent=0, spaceAfter=4))
        styles.add(ParagraphStyle(name="RevCode", fontName="DejaVu", fontSize=8.5, leading=12,
                                  backColor=colors.HexColor("#F1F5F9"), borderPadding=7,
                                  borderColor=colors.HexColor("#CBD5E1"), borderWidth=0.5))
        styles.add(ParagraphStyle(name="RevSmall", fontName="DejaVu", fontSize=8, leading=11,
                                  textColor=colors.HexColor("#64748B")))

        story = [Paragraph("Lesson revision notes", styles["RevTitle"]),
                 Paragraph(xml_escape(f"Lesson: {safe_stem.replace('_', ' ')}"), styles["RevSmall"]),
                 Spacer(1, 7)]

        def safe_para_html(node):
            # Markdown HTML is generated locally; retain only simple formatting tags.
            allowed = {"p", "b", "strong", "i", "em", "code", "a", "sup", "sub", "br", "img", "span"}
            for tag in list(node.find_all(True)):
                if tag.name not in allowed:
                    tag.unwrap()
                elif tag.name == "a":
                    tag.attrs = {"href": tag.get("href", "#")}
                elif tag.name == "img":
                    tag.attrs = {k: tag.get(k) for k in ("src", "width", "height") if tag.get(k) is not None}
                else:
                    tag.attrs = {}
            return str(node)

        def add_list(node, ordered=False):
            items=[]
            for li in node.find_all("li", recursive=False):
                # Keep all inline children (especially equation images) in one
                # Paragraph. Splitting <strong>, text and <img> into separate
                # Paragraphs caused missing equations and isolated punctuation.
                inline_parts=[]
                for child in li.children:
                    child_name = getattr(child, "name", None)
                    if child_name in ("ul", "ol"):
                        continue
                    if child_name == "p":
                        if child.get_text(strip=True) or child.find("img"):
                            inline_parts.extend(str(x) for x in child.contents)
                    elif isinstance(child, NavigableString):
                        if str(child).strip():
                            inline_parts.append(xml_escape(str(child)))
                    elif child_name:
                        if child.get_text(strip=True) or child.name == "img":
                            inline_parts.append(safe_para_html(child))
                combined = "".join(inline_parts).strip()
                inner=[Paragraph(combined, styles["RevBody"])] if combined else [Paragraph(" ", styles["RevBody"])]
                items.append(ListItem(inner, leftIndent=12))
            if ordered:
                story.append(ListFlowable(items, bulletType="1", start="1", leftIndent=18,
                                          bulletFontName="DejaVu", bulletFontSize=8, spaceAfter=6))
            else:
                story.append(ListFlowable(items, bulletType="bullet", bulletChar="•", leftIndent=18,
                                          bulletFontName="DejaVu", bulletFontSize=8, spaceAfter=6))

        for node in soup.contents:
            if isinstance(node, NavigableString):
                if str(node).strip(): story.append(Paragraph(xml_escape(str(node).strip()), styles["RevBody"]))
                continue
            if not getattr(node, "name", None): continue
            name=node.name.lower()
            text_value=node.get_text(" ", strip=True)
            display_match=re.fullmatch(r"REVISIONDISPLAYMATH\d+TOKEN", text_value)
            if display_match and text_value in display_tokens:
                image_path=render_equation(display_tokens[text_value], True)
                img=RLImage(str(image_path))
                max_w=A4[0]-38*mm
                max_h=42*mm
                # math_to_image rasterises at 240 dpi; ReportLab's natural
                # image size is in points, so convert pixels to points first.
                natural_scale=72/240
                natural_w=img.imageWidth*natural_scale
                natural_h=img.imageHeight*natural_scale
                scale=min(max_w/natural_w, max_h/natural_h, 1.0)
                img.drawWidth=natural_w*scale
                img.drawHeight=natural_h*scale
                img.hAlign="CENTER"
                story.extend([Spacer(1,4), img, Spacer(1,7)])
            elif name in ("h1", "h2", "h3", "h4"):
                if name == "h1" and text_value.strip().lower() == "lesson revision notes":
                    continue
                style=styles["RevH1"] if name in ("h1",) else styles["RevH2"] if name in ("h2",) else styles["RevH3"]
                story.append(Paragraph(safe_para_html(node), style))
            elif name == "p":
                story.append(Paragraph(safe_para_html(node), styles["RevBody"]))
            elif name == "ul": add_list(node, False)
            elif name == "ol": add_list(node, True)
            elif name == "pre":
                story.append(Preformatted(node.get_text(), styles["RevCode"], maxLineLength=100))
                story.append(Spacer(1,6))
            elif name == "blockquote":
                story.append(Paragraph(safe_para_html(node), ParagraphStyle(
                    "RevQuote", parent=styles["RevBody"], leftIndent=12,
                    borderColor=colors.HexColor("#94A3B8"), borderWidth=1, borderPadding=6)))
            elif name == "hr": story.append(HRFlowable(width="100%", color=colors.HexColor("#CBD5E1"), spaceBefore=5, spaceAfter=8))
            elif name == "table":
                rows=[]
                for tr in node.find_all("tr"):
                    cells=[]
                    for cell in tr.find_all(["th","td"], recursive=False):
                        cells.append(Paragraph(safe_para_html(cell), styles["RevBody"]))
                    if cells: rows.append(cells)
                if rows:
                    col_count=max(len(row) for row in rows)
                    for row in rows:
                        row.extend([Paragraph("", styles["RevBody"])]*(col_count-len(row)))
                    table=Table(rows, repeatRows=1, hAlign="LEFT")
                    table.setStyle(TableStyle([
                        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E8EEF7")),
                        ("TEXTCOLOR",(0,0),(-1,0),colors.HexColor("#17365D")),
                        ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CBD5E1")),
                        ("VALIGN",(0,0),(-1,-1),"TOP"),
                        ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
                        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
                    ]))
                    story.extend([table,Spacer(1,8)])

        if not story:
            story.append(Paragraph("No lesson notes were supplied.", styles["RevBody"]))

        output=BytesIO()
        doc=SimpleDocTemplate(output, pagesize=A4, rightMargin=19*mm, leftMargin=19*mm,
                              topMargin=18*mm, bottomMargin=18*mm,
                              title=safe_stem, author="Tutoring Automation")
        def page_decor(canvas, document):
            canvas.saveState()
            width,height=A4
            canvas.setStrokeColor(colors.HexColor("#D7DFEA"))
            canvas.setLineWidth(0.5)
            canvas.line(19*mm, 14*mm, width-19*mm, 14*mm)
            canvas.setFont("DejaVu",7.5)
            canvas.setFillColor(colors.HexColor("#64748B"))
            canvas.drawString(19*mm, 9*mm, "Student revision notes")
            canvas.drawRightString(width-19*mm, 9*mm, f"Page {document.page}")
            canvas.restoreState()
        doc.build(story, onFirstPage=page_decor, onLaterPages=page_decor)
        return output.getvalue()


@app.post("/render-summary-pdf")
def render_summary_pdf(request: RenderSummaryPdfRequest):
    filename = safe_filename(request.file_name)
    stem = Path(filename).stem
    pdf_filename = f"{stem}_revision_notes.pdf"
    pdf_bytes = build_revision_pdf(request.markdown_text, filename)
    pdf_item = upload_pdf_to_onedrive(
        pdf_filename, pdf_bytes, request.graph_access_token, request.folder_id
    )
    return {
        "pdf_file": pdf_item,
        "pdf_filename": pdf_filename,
        "pdf_size_bytes": len(pdf_bytes),
        "message": f"Rendered and saved {pdf_filename} to OneDrive.",
    }
