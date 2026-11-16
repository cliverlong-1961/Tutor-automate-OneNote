# User guide

## Sign in and choose a recording

1. Open the application and sign in with your Microsoft account.
2. Open the OneDrive panel.
3. Navigate using Root, folder breadcrumbs and Back.
4. Select an MP4 or MKV recording in the Video recording selection area.

The picker displays filename, format, size, modification date and processing status.

## Processing status

- **Not processed** — no matching transcript and summary were found.
- **Transcript only** — a transcript exists but no matching summary was found.
- **Completed** — both transcript and summary were found.

The app matches files by filename stem in the currently loaded folder:

| File | Example |
|---|---|
| Recording | `lesson-name.mp4` |
| Transcript | `lesson-name.txt` or `lesson-name_transcript.txt` |
| Summary | `lesson-name_summary.txt` |

The inspected lesson-grouping code recognises MP4 recordings. The recording picker also lists MKV files for processing.

## Transcribe and summarise

1. Select a recording.
2. Choose **Transcribe and summarise selected recording**.
3. The frontend obtains a Microsoft Graph access token.
4. The backend downloads the recording from OneDrive.
5. AssemblyAI transcribes it; MKV audio is first extracted using FFmpeg.
6. The transcript is returned to the frontend.
7. The frontend sends the transcript to `/summarise-and-save`.
8. The backend saves the transcript to OneDrive, asks Gemini to generate revision notes, and saves the summary.

The transcript is saved before summary generation. If Gemini fails, the transcript may still be available.

## Generate a missing summary

If a transcript exists but the summary is missing, select the recording and choose **Generate summary from existing transcript**. The app downloads the existing transcript and sends it to `/generate-missing-summary`. This does not transcribe the recording again.

## View an existing summary

Select a matching `_summary.txt` file and use the summary retrieval action. The app downloads it from OneDrive and displays it in the summary panel.

## Prepare a lesson for OneNote

A recording, transcript and summary are required.

1. Select the recognised lesson.
2. Choose **Prepare lesson for OneNote**.
3. The app retrieves the summary and asks the backend `/render-summary-pdf` endpoint to create a PDF in the current OneDrive folder.
4. The frontend creates view-sharing links for the recording and PDF.
5. Select a OneNote notebook.
6. Select a section.
7. Choose **Create OneNote lesson page**.

The page contains a **Lesson files** section with links to the recording and PDF, followed by **Student revision notes**. Microsoft Graph renders the PDF pages as images within the OneNote page. The transcript remains in OneDrive but is not linked on the page.

The app displays an **Open created OneNote page** link when Microsoft Graph returns the page URL.

## Reprocessing

The normal processing button is disabled for a completed lesson. A separate reprocessing path asks for confirmation before replacing existing transcript and summary outputs.

## Privacy

The app currently requests anonymous view links for the recording and PDF. Anyone who obtains those links may be able to view the files. Consider the sensitivity of the recording before preparing a lesson.
