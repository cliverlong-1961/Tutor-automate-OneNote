import { useEffect, useRef, useState } from "react";
import ReactMarkdown from "react-markdown";
import remarkMath from "remark-math";
import rehypeKatex from "rehype-katex";
import "katex/dist/katex.min.css";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { loginRequest } from "./authConfig";

interface DriveItem {
  id: string;
  name: string;
  size?: number;
  folder?: { childCount: number };
  file?: { mimeType: string };
  lastModifiedDateTime?: string;
  webUrl?: string;
}

interface SharingLink {
  id: string;
  name: string;
  url: string;
}

interface LessonGroup {
  key: string;
  title: string;
  recording: DriveItem;
  transcript?: DriveItem;
  summary?: DriveItem;
}

interface OneNoteNotebook {
  id: string;
  displayName: string;
  isDefault?: boolean;
  isShared?: boolean;
}

interface OneNoteSection {
  id: string;
  displayName: string;
  isDefault?: boolean;
}

type DiagnosticId = "backend" | "gemini" | "assemblyai" | "auth" | "onedrive" | "onenote" | "pdf";
type DiagnosticState = { status: "idle" | "checking" | "ok" | "warning" | "error"; message: string; checkedAt?: string; durationMs?: number };

const API_BASE = "https://tutoring-automation-api.graycliff-b3a723e4.uksouth.azurecontainerapps.io";
const DIAGNOSTIC_LABELS: Record<DiagnosticId, string> = {
  backend: "Azure backend",
  gemini: "Gemini",
  assemblyai: "AssemblyAI",
  auth: "Microsoft sign-in",
  onedrive: "OneDrive",
  onenote: "OneNote",
  pdf: "PDF generation",
};


// Convert common LaTeX notation to readable Unicode text for OneNote.
// The OneNote Graph API removes external CSS and supports only a subset of HTML,
// so KaTeX's browser markup cannot be embedded as-is. React continues to use KaTeX.
function latexToOneNoteText(input: string): string {
  let value = input.trim();
  const readGroup = (source: string, start: number): [string, number] => {
    if (source[start] !== "{") return [source[start] || "", start + 1];
    let depth = 0;
    for (let i = start; i < source.length; i++) {
      if (source[i] === "{") depth++;
      if (source[i] === "}") {
        depth--;
        if (depth === 0) return [source.slice(start + 1, i), i + 1];
      }
    }
    return [source.slice(start + 1), source.length];
  };
  // Resolve fractions from the inside out; use a solidus so the result remains editable.
  for (let pass = 0; pass < 8 && /\\frac\s*\{/.test(value); pass++) {
    const match = /\\frac\s*\{/.exec(value);
    if (!match || match.index === undefined) break;
    const start = match.index;
    const numeratorStart = start + match[0].length - 1;
    const [numerator, afterNumerator] = readGroup(value, numeratorStart);
    const denominatorOpen = value.indexOf("{", afterNumerator);
    if (denominatorOpen < 0) break;
    const [denominator, afterDenominator] = readGroup(value, denominatorOpen);
    value = value.slice(0, start) + `(${latexToOneNoteText(numerator)})/(${latexToOneNoteText(denominator)})` + value.slice(afterDenominator);
  }
  value = value
    .replace(/\\sqrt\s*\{([^{}]*)\}/g, "√($1)")
    .replace(/\\left|\\right/g, "")
    .replace(/\\,|\\;|\\!|\\quad|\\qquad/g, " ")
    .replace(/\\cdot/g, "·").replace(/\\times/g, "×")
    .replace(/\\leq|\\le/g, "≤").replace(/\\geq|\\ge/g, "≥")
    .replace(/\\neq|\\ne/g, "≠").replace(/\\pm/g, "±")
    .replace(/\\infty/g, "∞").replace(/\\pi/g, "π")
    .replace(/\\theta/g, "θ").replace(/\\alpha/g, "α")
    .replace(/\\beta/g, "β").replace(/\\gamma/g, "γ")
    .replace(/\\lambda/g, "λ").replace(/\\mu/g, "μ")
    .replace(/\\sigma/g, "σ").replace(/\\Delta/g, "Δ")
    .replace(/\\sin\b/g, "sin").replace(/\\cos\b/g, "cos")
    .replace(/\\tan\b/g, "tan").replace(/\\cot\b/g, "cot")
    .replace(/\\log\b/g, "log").replace(/\\ln\b/g, "ln")
    .replace(/\\cdots|\\ldots/g, "…")
    .replace(/\\text\s*\{([^{}]*)\}/g, "$1")
    .replace(/\\operatorname\s*\{([^{}]*)\}/g, "$1")
    .replace(/\^\{?([0-9+-])\}?/g, (_m, c: string) => (({"0":"⁰","1":"¹","2":"²","3":"³","4":"⁴","5":"⁵","6":"⁶","7":"⁷","8":"⁸","9":"⁹","+":"⁺","-":"⁻"} as Record<string, string>)[c] || `^${c}`))
    .replace(/_\{?([0-9+-])\}?/g, (_m, c: string) => (({"0":"₀","1":"₁","2":"₂","3":"₃","4":"₄","5":"₅","6":"₆","7":"₇","8":"₈","9":"₉","+":"₊","-":"₋"} as Record<string, string>)[c] || `_${c}`))
    .replace(/[{}]/g, "")
    .replace(/\\([A-Za-z]+)/g, "$1")
    .replace(/\s+/g, " ");
  return value;
}

function escapeOneNoteHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function markdownInlineToOneNoteHtml(source: string): string {
  const tokens: string[] = [];
  const hold = (html: string) => {
    const marker = `@@ONENOTETOKEN${tokens.length}@@`;
    tokens.push(html);
    return marker;
  };
  let text = source;
  // Protect math before escaping or interpreting Markdown punctuation.
  text = text.replace(/\\\((.+?)\\\)|\$([^$\n]+)\$/g, (_m, a: string, b: string) =>
    hold(`<span style="font-family:Cambria Math,serif">${escapeOneNoteHtml(latexToOneNoteText(a || b))}</span>`));
  text = text.replace(/`([^`]+)`/g, (_m, code: string) => hold(`<code>${escapeOneNoteHtml(code)}</code>`));
  text = escapeOneNoteHtml(text);
  text = text.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2">$1</a>');
  text = text.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/__(.+?)__/g, "<b>$1</b>");
  text = text.replace(/\*(.+?)\*/g, "<i>$1</i>").replace(/_(.+?)_/g, "<i>$1</i>");
  return text.replace(/@@ONENOTETOKEN(\d+)@@/g, (_m, n: string) => tokens[Number(n)] || "");
}

function markdownToOneNoteHtml(markdown: string): string {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  const out: string[] = [];
  let paragraph: string[] = [];
  let listType: "ul" | "ol" | null = null;
  let codeLines: string[] | null = null;

  const flushParagraph = () => {
    if (paragraph.length) {
      out.push(`<p>${markdownInlineToOneNoteHtml(paragraph.join(" "))}</p>`);
      paragraph = [];
    }
  };
  const closeList = () => {
    if (listType) { out.push(`</${listType}>`); listType = null; }
  };
  const isTableDivider = (line: string) => /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(line);
  const cells = (line: string) => line.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map(c => c.trim());

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (codeLines) {
      if (/^\s*```/.test(line)) {
        out.push(`<pre>${escapeOneNoteHtml(codeLines.join("\n"))}</pre>`);
        codeLines = null;
      } else codeLines.push(line);
      continue;
    }
    if (/^\s*```/.test(line)) { flushParagraph(); closeList(); codeLines = []; continue; }
    if (/^\s*\$\$\s*$/.test(line) || /^\s*\\\[\s*$/.test(line)) {
      flushParagraph(); closeList();
      const end = /^\s*\$\$/.test(line) ? /^\s*\$\$\s*$/ : /^\s*\\\]\s*$/;
      const math: string[] = [];
      i++;
      while (i < lines.length && !end.test(lines[i])) math.push(lines[i++]);
      out.push(`<p style="text-align:center;font-family:Cambria Math,serif;font-size:14pt"><b>${escapeOneNoteHtml(latexToOneNoteText(math.join(" ")))}</b></p>`);
      continue;
    }
    if (i + 1 < lines.length && line.includes("|") && isTableDivider(lines[i + 1])) {
      flushParagraph(); closeList();
      const headers = cells(line);
      out.push('<table border="1" style="border-collapse:collapse;width:100%"><tr>' + headers.map(h => `<td style="background-color:#e9e5fb"><b>${markdownInlineToOneNoteHtml(h)}</b></td>`).join("") + "</tr>");
      i += 2;
      while (i < lines.length && lines[i].includes("|") && lines[i].trim()) {
        const row = cells(lines[i]);
        out.push("<tr>" + headers.map((_h, j) => `<td>${markdownInlineToOneNoteHtml(row[j] || "")}</td>`).join("") + "</tr>");
        i++;
      }
      out.push("</table>"); i--; continue;
    }
    const heading = /^(#{1,6})\s+(.+)$/.exec(line);
    if (heading) {
      flushParagraph(); closeList();
      const level = Math.min(Number(heading[1].length) + 1, 6);
      out.push(`<h${level}>${markdownInlineToOneNoteHtml(heading[2])}</h${level}>`);
      continue;
    }
    const bullet = /^\s*[-*+]\s+(.+)$/.exec(line);
    const numbered = /^\s*\d+[.)]\s+(.+)$/.exec(line);
    if (bullet || numbered) {
      flushParagraph();
      const nextType = bullet ? "ul" : "ol";
      if (listType !== nextType) { closeList(); out.push(`<${nextType}>`); listType = nextType; }
      out.push(`<li>${markdownInlineToOneNoteHtml((bullet || numbered)![1])}</li>`);
      continue;
    }
    if (!line.trim()) { flushParagraph(); closeList(); continue; }
    if (/^\s*([-*_])\1\1+\s*$/.test(line)) { flushParagraph(); closeList(); out.push("<hr />"); continue; }
    closeList(); paragraph.push(line.trim());
  }
  if (codeLines) out.push(`<pre>${escapeOneNoteHtml(codeLines.join("\n"))}</pre>`);
  flushParagraph(); closeList();
  return out.join("\n");
}

function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [userDetails, setUserDetails] = useState("");
  const [message, setMessage] = useState("");
  const [driveItems, setDriveItems] = useState<DriveItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [folderLoaded, setFolderLoaded] = useState(false);
  const [folderStack, setFolderStack] = useState<{ id: string; name: string }[]>([]);
  const restoredFolderAccountRef = useRef("");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [sharingLinks, setSharingLinks] = useState<SharingLink[]>([]);
  const [creatingLinks, setCreatingLinks] = useState(false);
  const [summaryText, setSummaryText] = useState("");
  const [loadingSummary, setLoadingSummary] = useState(false);
  const [testingGemini, setTestingGemini] = useState(false);
  const [diagnostics, setDiagnostics] = useState<Record<DiagnosticId, DiagnosticState>>({
    backend: { status: "idle", message: "Not checked" },
    gemini: { status: "idle", message: "Not checked" },
    assemblyai: { status: "idle", message: "Not checked" },
    auth: { status: "idle", message: "Not checked" },
    onedrive: { status: "idle", message: "Not checked" },
    onenote: { status: "idle", message: "Not checked" },
    pdf: { status: "idle", message: "Not checked" },
  });
  const [runningAllDiagnostics, setRunningAllDiagnostics] = useState(false);
  const [notebooks, setNotebooks] = useState<OneNoteNotebook[]>([]);
  const [loadingNotebooks, setLoadingNotebooks] = useState(false);
  const [selectedNotebookId, setSelectedNotebookId] = useState("");
  const [sections, setSections] = useState<OneNoteSection[]>([]);
  const [loadingSections, setLoadingSections] = useState(false);
  const [selectedSectionId, setSelectedSectionId] = useState("");
  const [creatingLessonPage, setCreatingLessonPage] = useState(false);
  const [createdPageUrl, setCreatedPageUrl] = useState("");
  const [lessonPageCreated, setLessonPageCreated] = useState(false);
  const [preparingLesson, setPreparingLesson] = useState(false);
  const [selectedVideoId, setSelectedVideoId] = useState("");
  const [incompleteVideosExpanded, setIncompleteVideosExpanded] = useState(true);
  const [completedVideosExpanded, setCompletedVideosExpanded] = useState(false);
  const [statusMinimized, setStatusMinimized] = useState(false);
  const [transcribingVideo, setTranscribingVideo] = useState(false);
  const [videoTranscript, setVideoTranscript] = useState("");
  const [transcriptFileName, setTranscriptFileName] = useState("");

  const [expandedPanels, setExpandedPanels] = useState<Set<string>>(new Set());
  const [panelOrder, setPanelOrder] = useState<string[]>([
    "notebooks",
    "sections",
    "lessons",
    "onedrive",
    "links",
    "summary",
  ]);

  // Expand or collapse an individual panel.
  // Manually opening a panel does not close other open panels.
  function togglePanel(panelId: string) {
    setExpandedPanels((previous) => {
      const updated = new Set(previous);
      if (updated.has(panelId)) {
        updated.delete(panelId);
      } else {
        updated.add(panelId);
      }
      return updated;
    });
  }

  // Put newly loaded content first and automatically collapse older panels.
  function activatePanel(panelId: string) {
    setPanelOrder((previous) => [
      panelId,
      ...previous.filter((id) => id !== panelId),
    ]);
    setExpandedPanels(new Set([panelId]));
  }

  function renderPanelHeader(panelId: string, title: string) {
    const isExpanded = expandedPanels.has(panelId);
    return (
      <button
        type="button"
        onClick={() => togglePanel(panelId)}
        aria-expanded={isExpanded}
        style={{
          width: "100%",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "14px 18px",
          marginTop: "12px",
          border: expandedPanels.has(panelId)
            ? "1px solid #7664c5"
            : "1px solid #d7d2e8",
          borderRadius: "10px",
          backgroundColor: expandedPanels.has(panelId) ? "#e9e5fb" : "#f1eff8",
          color: "#29235f",
          fontSize: "19px",
          fontWeight: "700",
          textAlign: "left",
          cursor: "pointer",
          boxShadow: expandedPanels.has(panelId)
            ? "0 2px 7px rgba(65, 52, 130, 0.10)"
            : "none",
          transition: "background-color 160ms ease, border-color 160ms ease",
        }}
      >
        <span>{title}</span>
        <span>{isExpanded ? "▲" : "▼"}</span>
      </button>
    );
  }

  async function getAccessToken(): Promise<string> {
    const account = instance.getActiveAccount() || accounts[0];
    if (!account) {
      throw new Error("No active Microsoft account.");
    }

    try {
      const response = await instance.acquireTokenSilent({
        ...loginRequest,
        account,
      });
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await instance.acquireTokenPopup({
          ...loginRequest,
          account,
        });
        return response.accessToken;
      }
      throw error;
    }
  }

  // Force MSAL to obtain a fresh Microsoft Graph token when the backend
  // reports that the token used for a OneDrive download was rejected.
  async function getFreshAccessToken(): Promise<string> {
    const account = instance.getActiveAccount() || accounts[0];
    if (!account) throw new Error("No active Microsoft account.");
    try {
      const response = await instance.acquireTokenSilent({
        ...loginRequest,
        account,
        forceRefresh: true,
      });
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const response = await instance.acquireTokenPopup({
          ...loginRequest,
          account,
        });
        return response.accessToken;
      }
      throw error;
    }
  }

  async function signIn() {
    try {
      await instance.loginRedirect(loginRequest);
    } catch (error) {
      console.error(error);
      setMessage("Sign-in failed. Check the browser console.");
    }
  }

  async function signOut() {
    await instance.logoutRedirect();
  }

  async function testGraph() {
    try {
      const token = await getAccessToken();
      const response = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) {
        throw new Error(`Graph returned ${response.status}: ${await response.text()}`);
      }

      const data = await response.json();
      setUserDetails(
        `Name: ${data.displayName}\nEmail: ${data.mail || data.userPrincipalName}`
      );
      setMessage("Microsoft Graph connection successful.");
    } catch (error) {
      console.error(error);
      setMessage(error instanceof Error ? error.message : "Microsoft Graph test failed.");
    }
  }

  async function loadOneNoteNotebooks() {
    setLoadingNotebooks(true);
    setNotebooks([]);
    setMessage("Loading OneNote notebooks...");

    try {
      const token = await getAccessToken();
      const response = await fetch(
        "https://graph.microsoft.com/v1.0/me/onenote/notebooks",
        {
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json",
          },
        }
      );

      if (!response.ok) {
        throw new Error(
          `Unable to retrieve OneNote notebooks (${response.status}): ${await response.text()}`
        );
      }

      const data = await response.json();
      const loadedNotebooks: OneNoteNotebook[] = data.value || [];
      setNotebooks(loadedNotebooks);

      if (loadedNotebooks.length > 0) {
        const preferredNotebookId = localStorage.getItem(
          "tutoringAutomation.preferredNotebookId"
        );
        const preferredNotebook = loadedNotebooks.find(
          (notebook) => notebook.id === preferredNotebookId
        );

        if (preferredNotebook) {
          setMessage(`Restoring saved OneNote destination: ${preferredNotebook.displayName}...`);
          await loadOneNoteSections(preferredNotebook.id, true);
        } else {
          if (preferredNotebookId) {
            localStorage.removeItem("tutoringAutomation.preferredNotebookId");
            localStorage.removeItem("tutoringAutomation.preferredSectionId");
            setSelectedNotebookId("");
            setSelectedSectionId("");
          }
          activatePanel("notebooks");
          setMessage(`Successfully retrieved ${loadedNotebooks.length} notebooks. Select your preferred destination.`);
        }
      } else {
        setMessage("No OneNote notebooks were found.");
      }
    } catch (error) {
      console.error(error);
      setMessage(
        error instanceof Error ? error.message : "Unable to retrieve OneNote notebooks."
      );
    } finally {
      setLoadingNotebooks(false);
    }
  }

  async function loadOneNoteSections(
    notebookId: string,
    restoreSavedSection = false
  ) {
    setLoadingSections(true);
    setSections([]);
    setSelectedNotebookId(notebookId);
    setSelectedSectionId("");
    setCreatedPageUrl("");
    setLessonPageCreated(false);
    localStorage.setItem("tutoringAutomation.preferredNotebookId", notebookId);

    if (!restoreSavedSection) {
      localStorage.removeItem("tutoringAutomation.preferredSectionId");
    }

    setMessage("Loading OneNote sections...");

    try {
      const token = await getAccessToken();
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/onenote/notebooks/${notebookId}/sections`,
        {
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json",
          },
        }
      );

      if (!response.ok) {
        throw new Error(
          `Unable to retrieve OneNote sections (${response.status}): ${await response.text()}`
        );
      }

      const data = await response.json();
      const loadedSections: OneNoteSection[] = data.value || [];
      setSections(loadedSections);

      const savedSectionId = restoreSavedSection
        ? localStorage.getItem("tutoringAutomation.preferredSectionId")
        : null;
      const savedSection = savedSectionId
        ? loadedSections.find((section) => section.id === savedSectionId)
        : undefined;

      if (savedSection) {
        setSelectedSectionId(savedSection.id);
        activatePanel("createpage");
        setMessage(`Restored saved OneNote destination: ${savedSection.displayName}.`);
      } else {
        if (restoreSavedSection && savedSectionId) {
          localStorage.removeItem("tutoringAutomation.preferredSectionId");
        }
        setSelectedSectionId("");
        activatePanel("sections");
        setMessage(`Successfully retrieved ${loadedSections.length} sections. Select your preferred section.`);
      }
    } catch (error) {
      console.error(error);
      setMessage(
        error instanceof Error ? error.message : "Unable to retrieve OneNote sections."
      );
    } finally {
      setLoadingSections(false);
    }
  }

  function escapeHtml(value: string): string {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  async function createOneNotePage() {
    if (!selectedSectionId) {
      setMessage("Please select a OneNote section first.");
      return;
    }
    if (!summaryText.trim()) {
      setMessage("Please load the lesson summary from OneDrive first.");
      return;
    }
    if (sharingLinks.length === 0) {
      setMessage("Please generate the sharing links for the lesson recording and files first.");
      return;
    }

    setCreatingLessonPage(true);
    setCreatedPageUrl("");
    setMessage("Creating OneNote lesson page...");

    try {
      const token = await getAccessToken();
      const summaryFile = driveItems.find(
        (item) => item.name.toLowerCase().endsWith("_summary.txt") && selectedIds.has(item.id)
      );
      const lessonName = summaryFile
        ? summaryFile.name.replace(/_summary\.txt$/i, "")
        : new Date().toLocaleDateString("en-GB");
      const pageTitle = `Lesson: ${lessonName}`;
      const fileLinksHtml = sharingLinks
        .filter((link) => !/_transcript\.txt$|\.txt$/i.test(link.name))
        .map(
          (link) =>
            `<li><a href="${escapeHtml(link.url)}">${escapeHtml(link.name)}</a></li>`
        )
        .join("\n");

      // Find and download the PDF already saved in OneDrive. Microsoft Graph can
      // render a PDF data part as one image per PDF page in the new OneNote page.
      const pdfLink = sharingLinks.find((link) =>
        /_revision_notes\.pdf$/i.test(link.name)
      );
      if (!pdfLink) {
        throw new Error("The revision-notes PDF link is missing. Prepare the lesson again to generate its PDF and sharing link.");
      }

      setMessage("Downloading the revision-notes PDF and embedding its pages as images in OneNote...");
      const pdfResponse = await fetch(
        `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(pdfLink.id)}/content`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!pdfResponse.ok) {
        throw new Error(`Unable to download revision-notes PDF (${pdfResponse.status}): ${await pdfResponse.text()}`);
      }
      const pdfBlob = await pdfResponse.blob();
      if (!pdfBlob.size) {
        throw new Error("The revision-notes PDF downloaded from OneDrive is empty.");
      }

      // The PDF remains linked under Lesson files. The same PDF is also rendered
      // by OneNote as a sequence of page images beneath Student revision notes.
      const pageContent = `<!DOCTYPE html>
<html>
<head><title>${escapeOneNoteHtml(pageTitle)}</title><meta name="created" content="${new Date().toISOString()}" /></head>
<body>
  <h2>Lesson files</h2>
  <ul>${fileLinksHtml}</ul>
  <h2>Student revision notes</h2>
  <img data-render-src="name:revisionNotesPdf" alt="Student revision notes, rendered from the PDF" width="650" />
</body>
</html>`;

      // Let the browser create a standards-compliant multipart/form-data body.
      // The part name must match data-render-src="name:revisionNotesPdf".
      const multipartBody = new FormData();
      multipartBody.append(
        "Presentation",
        new Blob([pageContent], { type: "text/html" }),
        "presentation.html"
      );
      multipartBody.append(
        "revisionNotesPdf",
        new File([pdfBlob], "revision-notes.pdf", { type: "application/pdf" })
      );

      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/onenote/sections/${selectedSectionId}/pages`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json",
          },
          body: multipartBody,
        }
      );

      if (!response.ok) {
        throw new Error(
          `Unable to create OneNote lesson page (${response.status}): ${await response.text()}`
        );
      }

      const data = await response.json();
      const pageUrl = data.links?.oneNoteWebUrl?.href || data.webUrl || "";
      setCreatedPageUrl(pageUrl);
      setLessonPageCreated(true);
      setMessage(`Successfully created OneNote lesson page: ${data.title || pageTitle}`);
    } catch (error) {
      console.error(error);
      setMessage(
        error instanceof Error ? error.message : "Unable to create OneNote lesson page."
      );
    } finally {
      setCreatingLessonPage(false);
    }
  }

  async function loadFolder(folderId?: string) {
    setLoading(true);
    setMessage("");
    setDriveItems([]);
    setSelectedIds(new Set());
    setSharingLinks([]);
    setSummaryText("");
    setSelectedVideoId("");
    setVideoTranscript("");
    setTranscriptFileName("");

    try {
      const token = await getAccessToken();
      const endpoint = folderId
        ? `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(folderId)}/children`
        : "https://graph.microsoft.com/v1.0/me/drive/root/children";

      const response = await fetch(
        `${endpoint}?$select=id,name,folder,file,size,lastModifiedDateTime,webUrl&$top=100`,
        { headers: { Authorization: `Bearer ${token}` } }
      );

      if (!response.ok) {
        throw new Error(`Unable to load folder (${response.status}): ${await response.text()}`);
      }

      const data = await response.json();
      const items: DriveItem[] = data.value || [];
      setDriveItems(items);
      setFolderLoaded(true);
      activatePanel("onedrive");
      setMessage(`Folder loaded successfully. ${items.length} items found.`);
      return true;
    } catch (error) {
      console.error(error);
      setMessage(
        error instanceof Error
          ? error.message
          : "Unable to load folder. Check the browser console."
      );
      return false;
    } finally {
      setLoading(false);
    }
  }

  function getRecordingOutputStatus(recording: DriveItem): "Not processed" | "Transcript only" | "Completed" {
    const stem = recording.name.replace(/\.(mp4|mkv)$/i, "").toLowerCase();
    const hasTranscript = driveItems.some(
      (item) => !item.folder && [
        `${stem}.txt`,
        `${stem}_transcript.txt`,
      ].includes(item.name.toLowerCase())
    );
    const hasSummary = driveItems.some(
      (item) => !item.folder && item.name.toLowerCase() === `${stem}_summary.txt`
    );

    if (hasTranscript && hasSummary) return "Completed";
    if (hasTranscript) return "Transcript only";
    return "Not processed";
  }

  async function transcribeSelectedVideo(forceReprocess = false) {
    const selectedVideo = driveItems.find(
      (item) => item.id === selectedVideoId && !item.folder
    );
    if (!selectedVideo) {
      setMessage("Select an MP4 or MKV recording first.");
      return;
    }

    const outputStatus = getRecordingOutputStatus(selectedVideo);
    if (outputStatus === "Completed" && !forceReprocess) {
      setMessage(`${selectedVideo.name} has already been transcribed and summarised. Choose Reprocess recording only if you want to replace the existing files.`);
      return;
    }
    if (outputStatus !== "Not processed" && forceReprocess) {
      const confirmed = window.confirm(
        `This recording already has ${outputStatus === "Completed" ? "a transcript and summary" : "a transcript"}. Reprocessing will submit the video to AssemblyAI again and replace the existing transcript and summary in OneDrive. Continue?`
      );
      if (!confirmed) return;
    }

    setTranscribingVideo(true);
    setVideoTranscript("");
    setTranscriptFileName("");
    setMessage(`Sending ${selectedVideo.name} to AssemblyAI. This may take some time...`);

    try {
      let token = await getAccessToken();
      const transcriptionEndpoint = "https://tutoring-automation-api.graycliff-b3a723e4.uksouth.azurecontainerapps.io/transcribe-onedrive";

      const submitTranscription = (accessToken: string) => fetch(transcriptionEndpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          item_id: selectedVideo.id,
          file_name: selectedVideo.name,
          graph_access_token: accessToken,
        }),
      });

      let response = await submitTranscription(token);
      let data = await response.json().catch(() => ({}));

      // A stale delegated token can sometimes work for listing OneDrive but
      // fail when downloading file content. Refresh it once and retry.
      if (response.status === 401) {
        setMessage("Microsoft Graph rejected the download token. Refreshing your Microsoft sign-in and retrying once...");
        token = await getFreshAccessToken();
        response = await submitTranscription(token);
        data = await response.json().catch(() => ({}));
      }

      if (!response.ok) {
        const detail = data.detail;
        const message = typeof detail === "string"
          ? detail
          : detail?.message || JSON.stringify(detail || data) || `Transcription service returned ${response.status}.`;
        throw new Error(message);
      }
      const transcriptText = data.text || "";
      setVideoTranscript(transcriptText);
      setTranscriptFileName(selectedVideo.name);
      setMessage("Transcription completed. Saving the transcript, then generating student revision notes with Gemini...");

      const currentFolderId = folderStack.length > 0
        ? folderStack[folderStack.length - 1].id
        : null;
      const saveResponse = await fetch("https://tutoring-automation-api.graycliff-b3a723e4.uksouth.azurecontainerapps.io/summarise-and-save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          file_name: selectedVideo.name,
          transcript_text: transcriptText,
          graph_access_token: token,
          folder_id: currentFolderId,
        }),
      });
      const savedData = await saveResponse.json().catch(() => ({}));
      if (!saveResponse.ok) {
        throw new Error(savedData.detail || `Lesson summary service returned ${saveResponse.status}. The transcript is available above but has not been saved yet.`);
      }

      setSummaryText(savedData.summary || "");
      if (savedData.transcript_file && savedData.summary_file) {
        setDriveItems((previous) => {
          const additions = [savedData.transcript_file, savedData.summary_file] as DriveItem[];
          const addedIds = new Set(additions.map((item) => item.id));
          const addedNames = new Set(additions.map((item) => item.name.toLowerCase()));
          return [...previous.filter((item) => !addedIds.has(item.id) && !addedNames.has(item.name.toLowerCase())), ...additions];
        });
      }
      activatePanel("summary");
      setMessage(savedData.message || `Transcript and student revision summary saved to OneDrive for ${selectedVideo.name}.`);
    } catch (error) {
      console.error(error);
      const detail = error instanceof Error ? error.message : "Unable to transcribe recording.";
      setMessage(
        detail.includes("Failed to fetch")
          ? "Cannot reach the local transcription backend. Start it using Run-Backend.bat, then try again."
          : detail
      );
    } finally {
      setTranscribingVideo(false);
    }
  }

  function getSavedFolderStorageKey(): string | null {
    const account = instance.getActiveAccount() || accounts[0];
    if (!account) return null;
    const accountIdentifier = account.homeAccountId || account.username;
    return `tutoringAutomation.lastOneDriveFolder.${accountIdentifier}`;
  }

  function saveLastOneDriveFolder(stack: { id: string; name: string }[]) {
    const storageKey = getSavedFolderStorageKey();
    if (!storageKey) return;
    try {
      localStorage.setItem(storageKey, JSON.stringify({ folderStack: stack }));
    } catch (error) {
      console.warn("Unable to save the last OneDrive folder in this browser.", error);
    }
  }

  async function loadOneDrive() {
    const loaded = await loadFolder();
    if (loaded) {
      setFolderStack([]);
      saveLastOneDriveFolder([]);
    }
  }

  async function openFolder(item: DriveItem) {
    if (!item.folder) return;
    const newStack = [...folderStack, { id: item.id, name: item.name }];
    if (await loadFolder(item.id)) {
      setFolderStack(newStack);
      saveLastOneDriveFolder(newStack);
    }
  }

  async function goBack() {
    if (folderStack.length === 0) return;
    const newStack = folderStack.slice(0, -1);
    if (await loadFolder(newStack[newStack.length - 1]?.id)) {
      setFolderStack(newStack);
      saveLastOneDriveFolder(newStack);
    }
  }

  async function navigateToBreadcrumb(index: number) {
    const newStack = folderStack.slice(0, index + 1);
    if (await loadFolder(newStack[index].id)) {
      setFolderStack(newStack);
      saveLastOneDriveFolder(newStack);
    }
  }

  function toggleSelection(itemId: string) {
    setSelectedIds((previous) => {
      const updated = new Set(previous);
      if (updated.has(itemId)) updated.delete(itemId);
      else updated.add(itemId);
      return updated;
    });
  }

  async function createSharingLinks() {
    if (selectedIds.size === 0) {
      setMessage("Please select at least one file.");
      return;
    }

    setCreatingLinks(true);
    setSharingLinks([]);
    setMessage("Creating sharing links...");
    const results: SharingLink[] = [];

    try {
      const token = await getAccessToken();
      const selectedItems = driveItems.filter(
        (item) => selectedIds.has(item.id) && !item.folder
      );

      for (const item of selectedItems) {
        const response = await fetch(
          `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(item.id)}/createLink`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${token}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify({ type: "view", scope: "anonymous" }),
          }
        );

        if (!response.ok) {
          throw new Error(
            `Failed to create link for ${item.name} (${response.status}): ${await response.text()}`
          );
        }

        const data = await response.json();
        const url = data.link?.webUrl;
        if (!url) {
          throw new Error(`Microsoft Graph did not return a sharing URL for ${item.name}.`);
        }

        results.push({ id: item.id, name: item.name, url });
        setSharingLinks([...results]);
      }

      if (results.length > 0) {
        activatePanel("links");
      }
      setMessage(`Successfully created ${results.length} sharing links.`);
    } catch (error) {
      console.error(error);
      setMessage(
        error instanceof Error ? error.message : "Unable to create sharing links."
      );
    } finally {
      setCreatingLinks(false);
    }
  }

  function setDiagnosticResult(id: DiagnosticId, status: DiagnosticState["status"], message: string, startedAt: number) {
    setDiagnostics((previous) => ({
      ...previous,
      [id]: { status, message, checkedAt: new Date().toLocaleTimeString(), durationMs: Date.now() - startedAt },
    }));
  }

  async function runDiagnostic(id: DiagnosticId) {
    const startedAt = Date.now();
    setDiagnostics((previous) => ({ ...previous, [id]: { status: "checking", message: "Checking..." } }));
    const readJson = async (response: Response) => {
      const raw = await response.text();
      let data: any = {};
      try { data = raw ? JSON.parse(raw) : {}; } catch { data = { detail: raw }; }
      if (!response.ok) {
        const detail = data.detail || data.message || `HTTP ${response.status} ${response.statusText}`;
        const text = typeof detail === "string" ? detail : (detail.message || JSON.stringify(detail));
        throw new Error(`HTTP ${response.status}: ${text}`);
      }
      return data;
    };

    try {
      let result = "";
      if (id === "backend") {
        const data = await readJson(await fetch(`${API_BASE}/health`, { signal: AbortSignal.timeout(20000) }));
        result = `Healthy; model ${data.gemini_model || "not reported"}; ` +
          `Gemini key ${data.gemini_api_key_configured ? "configured" : "MISSING"}; ` +
          `AssemblyAI key ${data.assemblyai_key_configured ? "configured" : "MISSING"}.`;
        if (!data.gemini_api_key_configured || !data.assemblyai_key_configured) {
          setDiagnosticResult(id, "warning", result, startedAt); return;
        }
      } else if (id === "gemini") {
        const data = await readJson(await fetch(`${API_BASE}/test-gemini`, {
          method: "POST", headers: { "Content-Type": "application/json" }, signal: AbortSignal.timeout(210000),
        }));
        result = `Connected; model ${data.model || "configured model"}. ${data.message || "Test response received."}`;
      } else if (id === "assemblyai") {
        const data = await readJson(await fetch(`${API_BASE}/diagnostics/assemblyai`, { signal: AbortSignal.timeout(25000) }));
        result = data.message || "Authentication and API request succeeded.";
      } else if (id === "auth") {
        const token = await getAccessToken();
        const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName", {
          headers: { Authorization: `Bearer ${token}` }, signal: AbortSignal.timeout(20000),
        });const data = await readJson(response);
        result = `Authenticated as ${data.displayName || data.userPrincipalName || "Microsoft user"}; Graph token accepted.`;
      } else if (id === "onedrive") {
        const token = await getAccessToken();
        const response = await fetch("https://graph.microsoft.com/v1.0/me/drive/root?$select=id,name,webUrl", {
          headers: { Authorization: `Bearer ${token}` }, signal: AbortSignal.timeout(25000),
        });
        const data = await readJson(response);
        const remembered = folderStack.length ? ` Current folder: ${folderStack[folderStack.length - 1].name}.` : "";
        result = `OneDrive root accessible (${data.name || "root"}).${remembered}`;
      } else if (id === "onenote") {
        const token = await getAccessToken();
        const response = await fetch("https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=id,displayName", {
          headers: { Authorization: `Bearer ${token}`, Accept: "application/json" }, signal: AbortSignal.timeout(90000),
        });
        const data = await readJson(response);
        const count = Array.isArray(data.value) ? data.value.length : 0;
        if (!count) {
          result = "Graph responded, but no notebooks were returned.";
          setDiagnosticResult(id, "warning", result, startedAt); return;
        }
        let sectionNote = "";
        if (selectedNotebookId) {
          const sectionResponse = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/notebooks/${encodeURIComponent(selectedNotebookId)}/sections?$select=id,displayName`, {
            headers: { Authorization: `Bearer ${token}`, Accept: "application/json" }, signal: AbortSignal.timeout(45000),
          });
          const sectionData = await readJson(sectionResponse);
          sectionNote = ` Selected notebook sections: ${(sectionData.value || []).length}.`;
        }
        result = `Graph returned ${count} notebook(s).${sectionNote}`;
      } else if (id === "pdf") {
        const data = await readJson(await fetch(`${API_BASE}/diagnostics/pdf`, { signal: AbortSignal.timeout(20000) }));
        result = data.message || "PDF dependencies are available.";
      }
      setDiagnosticResult(id, "ok", result, startedAt);
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      const status = /HTTP 429|HTTP 503|HTTP 504|timed out|timeout/i.test(detail) ? "warning" : "error";
      setDiagnosticResult(id, status, detail.slice(0, 500), startedAt);
    }
  }

  async function runAllDiagnostics() {
    setRunningAllDiagnostics(true);
    const ids: DiagnosticId[] = ["backend", "gemini", "assemblyai", "auth", "onedrive", "onenote", "pdf"];
    try {
      for (const id of ids) await runDiagnostic(id);
      setMessage("Service diagnostics completed. Review the results in the Service status panel.");
    } finally {
      setRunningAllDiagnostics(false);
    }
  }

  async function testGeminiConnectivity() {
    setTestingGemini(true);
    setMessage("Testing Gemini connectivity. Waiting for the backend and Gemini response...");

    const endpoint = "https://tutoring-automation-api.graycliff-b3a723e4.uksouth.azurecontainerapps.io/test-gemini";
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), 210000);

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal: controller.signal,
      });
      const rawBody = await response.text();
      let data: any = {};
      try {
        data = rawBody ? JSON.parse(rawBody) : {};
      } catch {
        data = { detail: rawBody };
      }

      if (response.ok && data.ok) {
        setMessage(
          `Gemini connectivity successful. Model: ${data.model || "configured model"}. ` +
          `${data.message || "Gemini returned the expected test response."}`
        );
        return;
      }

      const detailValue = data.detail || data.message || rawBody || response.statusText || "No error details were returned.";
      const detail = typeof detailValue === "object"
        ? (detailValue.message || JSON.stringify(detailValue))
        : String(detailValue);
      if (response.status === 404) {
        setMessage(
          `Gemini connectivity test could not run — HTTP 404 Not Found. ` +
          `The Azure backend does not currently recognise /test-gemini. ` +
          `Deploy the backend version containing this endpoint. Server response: ${String(detail).slice(0, 1200)}`
        );
      } else {
        const upstreamCode = data.gemini_status_code ? ` Gemini status code: ${data.gemini_status_code}.` : "";
        const errorType = data.error_type ? ` Error type: ${data.error_type}.` : "";
        setMessage(
          `Gemini connectivity test failed. Backend HTTP status: ${response.status} ${response.statusText || ""}.` +
          `${upstreamCode}${errorType} ${detail.slice(0, 1500)}`
        );
      }
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        setMessage(
          "Gemini connectivity test timed out after 210 seconds. No response was received from the backend/Gemini within that period. " +
          "This does not prove Gemini is unavailable; check the Azure Container App logs and try again."
        );
      } else if (error instanceof TypeError) {
        setMessage(
          `Could not contact the Azure backend for the Gemini test. This is usually a network, DNS, TLS or CORS problem. Details: ${error.message}`
        );
      } else {
        setMessage(`Gemini connectivity test could not be completed: ${error instanceof Error ? error.message : String(error)}`);
      }
      console.error("Gemini connectivity test error:", error);
    } finally {
      window.clearTimeout(timeoutId);
      setTestingGemini(false);
    }
  }

  async function generateMissingSummary(lesson: LessonGroup) {
    if (!lesson.transcript || lesson.summary) return;

    setLoadingSummary(true);
    setMessage(`Asking Azure to retrieve the existing transcript for ${lesson.title}...`);
    try {
      let token = await getAccessToken();
      const currentFolderId = folderStack.length > 0
        ? folderStack[folderStack.length - 1].id
        : null;
      const endpoint = "https://tutoring-automation-api.graycliff-b3a723e4.uksouth.azurecontainerapps.io/generate-missing-summary";

      // This API expects the transcript contents, not just the OneDrive item ID.
      // Download the existing transcript from Graph and pass it as transcript_text.
      const downloadTranscript = (accessToken: string) => fetch(
        `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(lesson.transcript!.id)}/content`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );

      setMessage(`Downloading the existing transcript for ${lesson.title} from OneDrive...`);
      let transcriptResponse = await downloadTranscript(token);
      if (transcriptResponse.status === 401) {
        setMessage("Microsoft Graph rejected the download token. Refreshing your Microsoft sign-in and retrying once...");
        token = await getFreshAccessToken();
        transcriptResponse = await downloadTranscript(token);
      }
      if (!transcriptResponse.ok) {
        throw new Error(`Unable to download the existing transcript (${transcriptResponse.status}): ${await transcriptResponse.text()}`);
      }
      const transcriptText = await transcriptResponse.text();
      if (!transcriptText.trim()) {
        throw new Error("The existing transcript file is empty. No summary was generated.");
      }

      const sendRequest = (accessToken: string) => fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          file_name: lesson.recording.name,
          transcript_text: transcriptText,
          graph_access_token: accessToken,
          folder_id: currentFolderId,
        }),
      });

      setMessage(`Sending the existing transcript for ${lesson.title} to Azure to generate its summary...`);
      let response = await sendRequest(token);
      if (response.status === 401) {
        setMessage("The summary service rejected the Microsoft token. Refreshing your Microsoft sign-in and retrying once...");
        token = await getFreshAccessToken();
        response = await sendRequest(token);
      }

      const rawBody = await response.text();
      let data: any = {};
      try {
        data = rawBody ? JSON.parse(rawBody) : {};
      } catch {
        data = { detail: rawBody };
      }
      if (!response.ok) {
        const detail = data.detail || data.message || `Summary service returned HTTP ${response.status}.`;
        throw new Error(typeof detail === "object" ? (detail.message || JSON.stringify(detail)) : String(detail));
      }

      setSummaryText(data.summary || "");
      if (data.summary_file) {
        const added = data.summary_file as DriveItem;
        setDriveItems((previous) => [
          ...previous.filter((item) => item.id !== added.id && item.name.toLowerCase() !== added.name.toLowerCase()),
          added,
        ]);
      }
      setMessage(data.message || `Missing summary generated for ${lesson.title}.`);
    } catch (error) {
      console.error(error);
      setMessage(error instanceof Error ? error.message : "Unable to generate the missing summary.");
    } finally {
      setLoadingSummary(false);
    }
  }

  async function prepareLessonForOneNote() {
    const selectedLesson = lessonGroups.find((lesson) =>
      selectedIds.has(lesson.recording.id)
    );

    if (!selectedLesson) {
      setMessage("Select a recognised lesson first.");
      return;
    }
    if (!selectedLesson.transcript || !selectedLesson.summary) {
      setMessage("This lesson is incomplete. A recording, transcript and summary are required.");
      return;
    }

    setPreparingLesson(true);
    setSharingLinks([]);
    setSummaryText("");
    setCreatedPageUrl("");
    setLessonPageCreated(false);
    setMessage("Preparing lesson: rendering the revision-notes PDF and creating sharing links...");

    try {
      const token = await getAccessToken();
      const currentFolderId = folderStack.length > 0
        ? folderStack[folderStack.length - 1].id
        : null;
      const summaryResponse = await fetch(
        `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(selectedLesson.summary.id)}/content`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!summaryResponse.ok) {
        throw new Error(`Unable to retrieve summary (${summaryResponse.status}): ${await summaryResponse.text()}`);
      }
      const summary = await summaryResponse.text();
      setSummaryText(summary);
      setMessage(`Rendering mathematical notes to PDF for ${selectedLesson.title}...`);

      const pdfResponse = await fetch("https://tutoring-automation-api.graycliff-b3a723e4.uksouth.azurecontainerapps.io/render-summary-pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          file_name: selectedLesson.recording.name,
          markdown_text: summary,
          graph_access_token: token,
          folder_id: currentFolderId || null,
        }),
      });
      const pdfData = await pdfResponse.json();
      if (!pdfResponse.ok) {
        throw new Error(pdfData.detail || `PDF generation failed (${pdfResponse.status}).`);
      }
      const pdfFile = pdfData.pdf_file as DriveItem;
      if (!pdfFile?.id || !pdfFile?.name) {
        throw new Error("PDF service did not return the saved OneDrive PDF file.");
      }
      setDriveItems((current) => [
        ...current.filter((item) => item.id !== pdfFile.id && item.name !== pdfFile.name),
        pdfFile,
      ]);

      const files = [selectedLesson.recording, pdfFile];
      const links: SharingLink[] = [];
      for (const item of files) {
        const response = await fetch(
          `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(item.id)}/createLink`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${token}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify({ type: "view", scope: "anonymous" }),
          }
        );
        if (!response.ok) {
          throw new Error(`Failed to create link for ${item.name} (${response.status}): ${await response.text()}`);
        }
        const data = await response.json();
        const url = data.link?.webUrl;
        if (!url) throw new Error(`Microsoft Graph did not return a sharing URL for ${item.name}.`);
        links.push({ id: item.id, name: item.name, url });
      }
      // Commit both links together so the workflow status sees the complete pair.
      setSharingLinks(links);
      setMessage(`Lesson prepared: ${selectedLesson.title}. PDF revision notes saved to OneDrive; sharing links created for the recording and PDF.`);
    } catch (error) {
      console.error(error);
      setMessage(error instanceof Error ? error.message : "Unable to prepare lesson.");
    } finally {
      setPreparingLesson(false);
    }
  }

  async function loadLessonSummary() {
    const summaryFile = driveItems.find(
      (item) =>
        selectedIds.has(item.id) &&
        !item.folder &&
        item.name.toLowerCase().endsWith("_summary.txt")
    );

    if (!summaryFile) {
      setMessage("The selected lesson has no matching summary TXT file. Check the Recognised lessons panel.");
      return;
    }

    setLoadingSummary(true);
    setSummaryText("");
    setMessage("Retrieving lesson summary...");

    try {
      const token = await getAccessToken();
      const response = await fetch(
        `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(summaryFile.id)}/content`,
        { headers: { Authorization: `Bearer ${token}` } }
      );

      if (!response.ok) {
        throw new Error(
          `Unable to retrieve summary (${response.status}): ${await response.text()}`
        );
      }

      const text = await response.text();
      setSummaryText(text);
      activatePanel("summary");
      setMessage(`Summary retrieved successfully: ${summaryFile.name}`);
    } catch (error) {
      console.error(error);
      setMessage(
        error instanceof Error ? error.message : "Unable to retrieve lesson summary."
      );
    } finally {
      setLoadingSummary(false);
    }
  }

  function formatSize(bytes?: number) {
    if (bytes === undefined) return "-";
    if (bytes < 1024) return `${bytes} bytes`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  }

  const summaryFileSelected = driveItems.some(
    (item) =>
      selectedIds.has(item.id) &&
      !item.folder &&
      item.name.toLowerCase().endsWith("_summary.txt")
  );

  // Recognise lesson recordings and match their transcript and summary files
  // by filename. Expected pattern:
  // recording: 2026-11-10 10-00-00.mp4
  // transcript: 2026-11-10 10-00-00.txt (or *_transcript.txt)
  // summary:    2026-11-10 10-00-00_summary.txt
  const lessonGroups: LessonGroup[] = driveItems
    .filter((item) => !item.folder && /\.mp4$/i.test(item.name))
    .map((recording) => {
      const stem = recording.name.replace(/\.mp4$/i, "");
      const transcript = driveItems.find(
        (item) =>
          !item.folder &&
          (item.name.toLowerCase() === `${stem}.txt`.toLowerCase() ||
            item.name.toLowerCase() === `${stem}_transcript.txt`.toLowerCase())
      );
      const summary = driveItems.find(
        (item) =>
          !item.folder &&
          item.name.toLowerCase() === `${stem}_summary.txt`.toLowerCase()
      );

      return {
        key: recording.id,
        title: stem,
        recording,
        transcript,
        summary,
      };
    });

  function selectLesson(lesson: LessonGroup) {
    const ids = [lesson.recording.id, lesson.transcript?.id, lesson.summary?.id]
      .filter((id): id is string => Boolean(id));
    setSelectedIds(new Set(ids));
    setSharingLinks([]);
    setSummaryText("");
    setCreatedPageUrl("");
    setMessage(
      `Selected lesson ${lesson.title}. ${
        ids.length
      } of 3 expected files found. Generate sharing links and load the summary.`
    );
  }

  // Panel contents are stored by ID; panelOrder controls their displayed order.
  const selectedLessonGroup = lessonGroups.find((lesson) =>
    selectedIds.has(lesson.recording.id)
  );
  // Preparation now creates two links only: the recording and the PDF.
  // Confirm both expected files are represented in state before marking complete.
  const lessonPrepared = Boolean(
    selectedLessonGroup &&
    summaryText.trim() &&
    sharingLinks.some((link) => link.id === selectedLessonGroup.recording.id) &&
    sharingLinks.some((link) => /_revision_notes\.pdf$/i.test(link.name))
  );
  const notebookSelected = Boolean(selectedNotebookId);
  const sectionSelected = Boolean(selectedSectionId);
  const pageCreated = lessonPageCreated;

  function statusLine(label: string, done: boolean) {
    return (
      <div className={done ? "workflow-status done" : "workflow-status pending"}>
        <span>{done ? "✓" : "○"}</span>
        <span>{label}</span>
        <strong>{done ? "Complete" : "Pending"}</strong>
      </div>
    );
  }

  const panelContent: Record<string, React.ReactNode> = {
    notebooks: notebooks.length > 0 ? (
      <div>
        {renderPanelHeader("notebooks", `OneNote notebooks (${notebooks.length})`)}
        {expandedPanels.has("notebooks") && (
          <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left" }}>
            <thead>
              <tr>
                <th style={{ padding: "8px" }}>Notebook name</th>
                <th style={{ padding: "8px" }}>Default</th>
                <th style={{ padding: "8px" }}>Shared</th>
                <th style={{ padding: "8px" }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {notebooks.map((notebook) => (
                <tr key={notebook.id}>
                  <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                    {notebook.displayName}
                  </td>
                  <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                    {notebook.isDefault ? "Yes" : "No"}
                  </td>
                  <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                    {notebook.isShared ? "Yes" : "No"}
                  </td>
                  <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                    <button
                      onClick={() => loadOneNoteSections(notebook.id)}
                      disabled={loadingSections}
                    >
                      {selectedNotebookId === notebook.id ? "Selected" : "Select notebook"}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    ) : null,

    sections: sections.length > 0 ? (
      <div>
        {renderPanelHeader(
          "sections",
          `Sections in selected notebook (${
            notebooks.find((n) => n.id === selectedNotebookId)?.displayName || ""
          })`
        )}
        {expandedPanels.has("sections") && (
          <>
            <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left" }}>
              <thead>
                <tr>
                  <th style={{ padding: "8px" }}>Section name</th>
                  <th style={{ padding: "8px" }}>Default</th>
                  <th style={{ padding: "8px" }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {sections.map((section) => (
                  <tr key={section.id}>
                    <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                      {section.displayName}
                    </td>
                    <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                      {section.isDefault ? "Yes" : "No"}
                    </td>
                    <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                      <button
                        onClick={() => {
                          setSelectedSectionId(section.id);
                          localStorage.setItem("tutoringAutomation.preferredSectionId", section.id);
                          setCreatedPageUrl("");
                          setLessonPageCreated(false);
                          activatePanel("createpage");
                        }}
                      >
                        {selectedSectionId === section.id ? "Selected" : "Select section"}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

          </>
        )}
      </div>
    ) : null,

    createpage: selectedSectionId ? (
      <div>
        {renderPanelHeader("createpage", "Create lesson page")}
        {expandedPanels.has("createpage") && (
          <div style={{ padding: "10px 4px" }}>
            <p>
              Selected section:{" "}
              <strong>
                {sections.find((section) => section.id === selectedSectionId)?.displayName || "Selected section"}
              </strong>
            </p>
            <p style={{ color: "#666", fontSize: "13px" }}>
              The page will include links to the recording and typeset revision-notes PDF; the transcript is retained in OneDrive but not linked on the page.
            </p>
            <button
              onClick={createOneNotePage}
              disabled={creatingLessonPage || !summaryText.trim() || sharingLinks.length === 0 || lessonPageCreated}
            >
              {creatingLessonPage
                ? "Creating lesson page..."
                : lessonPageCreated
                  ? "Lesson page created"
                  : "Create OneNote lesson page"}
            </button>
            {createdPageUrl && (
              <p style={{ marginTop: "10px" }}>
                <a href={createdPageUrl} target="_blank" rel="noopener noreferrer">
                  Open created OneNote page
                </a>
              </p>
            )}
          </div>
        )}
      </div>
    ) : null,

    onedrive: folderLoaded ? (
      <div>
        {renderPanelHeader("onedrive", `OneDrive (${driveItems.length} items)`)}
        {expandedPanels.has("onedrive") && (
          <>
            <p>
              <button onClick={loadOneDrive}>Root</button>
              {folderStack.map((folder, index) => (
                <span key={folder.id}>
                  {" / "}
                  <button onClick={() => navigateToBreadcrumb(index)}>
                    {folder.name}
                  </button>
                </span>
              ))}
              {folderStack.length > 0 && (
                <button onClick={goBack} style={{ marginLeft: "15px" }}>
                  Back
                </button>
              )}
            </p>

            <div style={{ margin: "15px 0" }}>
              <strong>{selectedIds.size} files selected</strong>
              <span style={{ marginLeft: "18px", color: "#666", fontSize: "13px" }}>
                Use “Prepare lesson for OneNote” above to generate the revision-notes PDF and create sharing links.
              </span>
            </div>

            <div style={{
              margin: "15px 0",
              padding: "12px",
              border: "1px solid #e3dfed",
              borderRadius: "10px",
              background: "#f8f7fc",
            }}>
              <h3 style={{ margin: "0 0 8px", textAlign: "center", color: "#514879" }}>
                Video recording selection
              </h3>
              <p style={{ margin: "4px 0 10px", color: "#666", textAlign: "center" }}>
                Select an MP4 or MKV recording to transcribe with AssemblyAI.
              </p>
              {selectedVideoId && (() => {
                const selected = driveItems.find(
                  (item) => item.id === selectedVideoId && !item.folder
                );
                const matchingLesson = lessonGroups.find(
                  (lesson) => lesson.recording.id === selectedVideoId
                );
                return selected &&
                  getRecordingOutputStatus(selected) === "Transcript only" &&
                  matchingLesson?.transcript ? (
                    <div style={{ textAlign: "center", margin: "8px 0" }}>
                      <button
                        onClick={() => generateMissingSummary(matchingLesson)}
                        disabled={loadingSummary || transcribingVideo || testingGemini}
                        style={{ padding: "9px 16px", borderColor: "#b7791f", color: "#805400" }}
                      >
                        {loadingSummary ? "Generating summary..." : "Generate summary from existing transcript"}
                      </button>
                    </div>
                  ) : null;
              })()}
              <button
                onClick={() => transcribeSelectedVideo(false)}
                disabled={!selectedVideoId || transcribingVideo || testingGemini || Boolean(driveItems.find((item) => item.id === selectedVideoId && getRecordingOutputStatus(item) === "Completed"))}
                style={{ display: "block", margin: "8px auto 14px", padding: "9px 16px" }}
              >
                {transcribingVideo ? "Transcribing and summarising..." : "Transcribe and summarise selected recording"}
              </button>

              {(() => {
                const videoFiles = driveItems.filter(
                  (item) => !item.folder && /\.(mp4|mkv)$/i.test(item.name)
                );
                const incompleteVideos = videoFiles.filter(
                  (item) => getRecordingOutputStatus(item) !== "Completed"
                );
                const completedVideos = videoFiles.filter(
                  (item) => getRecordingOutputStatus(item) === "Completed"
                );
                const renderVideoFile = (item: DriveItem) => (
                  <label key={item.id} style={{
                    display: "flex", alignItems: "flex-start", gap: "9px",
                    padding: "8px", margin: "6px 0", borderRadius: "7px",
                    border: selectedVideoId === item.id ? "1px solid #7664c5" : "1px solid #ddd",
                    background: selectedVideoId === item.id ? "#eeebff" : "#fff",
                    cursor: "pointer",
                  }}>
                    <input
                      type="radio"
                      name="selectedVideo"
                      checked={selectedVideoId === item.id}
                      onChange={() => setSelectedVideoId(item.id)}
                      style={{ marginTop: "3px", accentColor: "#5145b5" }}
                    />
                    <span style={{ minWidth: 0, overflowWrap: "anywhere", flex: 1 }}>
                      <strong>{item.name}</strong>
                      <span style={{ display: "block", color: "#777", fontSize: "12px" }}>
                        {item.name.toLowerCase().endsWith(".mkv") ? "MKV" : "MP4"} · {formatSize(item.size)} · {item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleDateString("en-GB") : "Date unavailable"}
                      </span>
                      <span style={{
                        display: "inline-block", marginTop: "4px", padding: "2px 7px",
                        borderRadius: "10px", fontSize: "11px", fontWeight: 700,
                        color: getRecordingOutputStatus(item) === "Completed" ? "#176b35" : getRecordingOutputStatus(item) === "Transcript only" ? "#8a5700" : "#555",
                        background: getRecordingOutputStatus(item) === "Completed" ? "#e5f5e9" : getRecordingOutputStatus(item) === "Transcript only" ? "#fff2d6" : "#eeeeF2",
                      }}>
                        {getRecordingOutputStatus(item)}
                      </span>
                    </span>
                  </label>
                );

                return (
                  <>
                    <div style={{ border: "1px solid #d8d4e7", borderRadius: "8px", background: "#fff", margin: "8px 0" }}>
                      <button
                        type="button"
                        onClick={() => setIncompleteVideosExpanded((expanded) => !expanded)}
                        aria-expanded={incompleteVideosExpanded}
                        style={{ width: "100%", textAlign: "left", padding: "10px 12px", border: 0, background: "#eeebff", color: "#40366f", borderRadius: "7px", fontWeight: 700 }}
                      >
                        {incompleteVideosExpanded ? "▾" : "▸"} Incomplete / not processed recordings ({incompleteVideos.length})
                      </button>
                      {incompleteVideosExpanded && (
                        <div style={{ padding: "4px 10px 8px" }}>
                          {incompleteVideos.length === 0
                            ? <p style={{ color: "#666", margin: "8px 0" }}>There are no incomplete recordings in this folder.</p>
                            : incompleteVideos.map(renderVideoFile)}
                        </div>
                      )}
                    </div>

                    <div style={{ border: "1px solid #d8d4e7", borderRadius: "8px", background: "#fff", margin: "8px 0" }}>
                      <button
                        type="button"
                        onClick={() => setCompletedVideosExpanded((expanded) => !expanded)}
                        aria-expanded={completedVideosExpanded}
                        style={{ width: "100%", textAlign: "left", padding: "10px 12px", border: 0, background: "#e5f5e9", color: "#176b35", borderRadius: "7px", fontWeight: 700 }}
                      >
                        {completedVideosExpanded ? "▾" : "▸"} Completed recordings ({completedVideos.length})
                      </button>
                      {completedVideosExpanded && (
                        <div style={{ padding: "4px 10px 8px" }}>
                          {selectedVideoId && (() => {
                            const selected = driveItems.find(
                              (item) => item.id === selectedVideoId && !item.folder
                            );
                            return selected && completedVideos.some(
                              (video) => video.id === selected.id
                            ) ? (
                              <button
                                type="button"
                                onClick={() => transcribeSelectedVideo(true)}
                                disabled={transcribingVideo}
                                style={{
                                  display: "block",
                                  margin: "8px auto 10px",
                                  padding: "8px 14px",
                                  borderColor: "#b7791f",
                                  color: "#805400",
                                }}
                              >
                                Reprocess recording (replace existing files)
                              </button>
                            ) : null;
                          })()}
                          {completedVideos.length === 0
                            ? <p style={{ color: "#666", margin: "8px 0" }}>There are no completed recordings in this folder.</p>
                            : completedVideos.map(renderVideoFile)}
                        </div>
                      )}
                    </div>
                    {videoFiles.length === 0 && (
                      <p>No MP4 or MKV recordings in this folder.</p>
                    )}
                  </>
                );
              })()}

              <p style={{ margin: "8px 0 0", color: "#666", fontSize: "12px", textAlign: "center" }}>
                Completed recordings are protected from accidental retranscription. Reprocessing requires confirmation and replaces the existing transcript and summary.
              </p>
              {videoTranscript && (
                <div style={{ marginTop: "12px", padding: "10px", background: "#fff", border: "1px solid #d8d4e7", borderRadius: "7px" }}>
                  <strong>Transcript: {transcriptFileName}</strong>
                  <pre style={{ whiteSpace: "pre-wrap", overflowWrap: "anywhere", fontFamily: "Arial", fontSize: "13px", lineHeight: 1.5 }}>
                    {videoTranscript}
                  </pre>
                </div>
              )}
            </div>

            <details
              style={{
                margin: "15px 0",
                border: "1px solid #e3dfed",
                borderRadius: "10px",
                background: "#f8f7fc",
                overflow: "hidden",
              }}
            >
              <summary
                style={{
                  padding: "11px 13px",
                  cursor: "pointer",
                  fontWeight: 700,
                  color: "#514879",
                  background: "#f1eff8",
                  userSelect: "none",
                }}
              >
                Folder file list ({driveItems.length} items)
              </summary>
              <div style={{ padding: "8px 10px", overflowX: "auto" }}>
            {driveItems.length === 0 ? (
              <p>This folder is empty.</p>
            ) : (
              <table style={{ width: "100%", maxWidth: "100%", borderCollapse: "collapse", tableLayout: "fixed", textAlign: "left" }}>
                <colgroup>
                  <col style={{ width: "55px" }} />
                  <col />
                  <col style={{ width: "105px" }} />
                  <col style={{ width: "70px" }} />
                  <col style={{ width: "105px" }} />
                </colgroup>
                <thead>
                  <tr>
                    <th style={{ padding: "5px", width: "55px", minWidth: "55px", maxWidth: "55px" }}>Select</th>
                    <th style={{ padding: "5px" }}>Name</th>
                    <th style={{ padding: "8px" }}>Type</th>
                    <th style={{ padding: "8px" }}>Size</th>
                    <th style={{ padding: "8px" }}>Last modified</th>
                  </tr>
                </thead>
                <tbody>
                  {driveItems.map((item) => (
                    <tr key={item.id}>
                      <td style={{ padding: "5px", width: "55px", minWidth: "55px", maxWidth: "55px", borderTop: "1px solid #ddd", textAlign: "center", verticalAlign: "middle" }}>
                        <input
                          type="checkbox"
                          aria-label={item.folder ? `${item.name} is a folder and cannot be selected` : `Select ${item.name}`}
                          disabled={Boolean(item.folder)}
                          checked={!item.folder && selectedIds.has(item.id)}
                          onChange={() => {
                            if (!item.folder) toggleSelection(item.id);
                          }}
                          style={{
                            appearance: "auto",
                            WebkitAppearance: "checkbox",
                            display: "inline-block",
                            visibility: "visible",
                            opacity: 1,
                            width: "15px",
                            height: "15px",
                            minWidth: "15px",
                            margin: 0,
                            accentColor: "#5145b5",
                            cursor: item.folder ? "not-allowed" : "pointer",
                          }}
                        />
                      </td>
                      <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                        {item.folder ? (
                          <button
                            onClick={() => openFolder(item)}
                            style={{
                              border: "none",
                              background: "none",
                              cursor: "pointer",
                              fontSize: "inherit",
                              color: "#1565c0",
                            }}
                          >
                            📁 {item.name}
                          </button>
                        ) : (
                          <>📄 {item.name}</>
                        )}
                      </td>
                      <td style={{ padding: "8px" }}>
                        {item.folder
                          ? `Folder (${item.folder.childCount} items)`
                          : item.file?.mimeType || "File"}
                      </td>
                      <td style={{ padding: "8px" }}>{formatSize(item.size)}</td>
                      <td style={{ padding: "8px" }}>
                        {item.lastModifiedDateTime
                          ? new Date(item.lastModifiedDateTime).toLocaleDateString("en-GB")
                          : "-"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
              </div>
            </details>
          </>
        )}
      </div>
    ) : null,

    lessons: folderLoaded ? (
      <div>
        {renderPanelHeader(
          "lessons",
          `Recognised lessons (${lessonGroups.length})`
        )}
        {expandedPanels.has("lessons") && (
          <>
            <div className="workflow-box">
              <h3>OneDrive workflow</h3>
              {statusLine("Lesson recording selected", Boolean(selectedLessonGroup))}
              {statusLine("Matching transcript and summary available", Boolean(selectedLessonGroup?.transcript && selectedLessonGroup?.summary))}
              {statusLine("PDF generated and sharing links created", lessonPrepared)}
              <button
                onClick={prepareLessonForOneNote}
                disabled={!selectedLessonGroup || !selectedLessonGroup.transcript || !selectedLessonGroup.summary || preparingLesson}
                style={{ marginTop: "10px" }}
              >
                {preparingLesson ? "Preparing lesson..." : "Prepare lesson for OneNote"}
              </button>
            </div>
            {lessonGroups.length === 0 ? (
              <p>
                No MP4 lesson recordings were found in this folder. Open the folder
                containing the recordings and their matching TXT files.
              </p>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left" }}>
                <thead>
                  <tr>
                    <th style={{ padding: "8px" }}>Lesson</th>
                    <th style={{ padding: "8px" }}>Recording</th>
                    <th style={{ padding: "8px" }}>Transcript</th>
                    <th style={{ padding: "8px" }}>Summary</th>
                    <th style={{ padding: "8px" }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {lessonGroups.map((lesson) => {
                    const complete = Boolean(lesson.recording && lesson.transcript && lesson.summary);
                    const selected = selectedIds.has(lesson.recording.id);
                    return (
                      <tr key={lesson.key}>
                        <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                          {lesson.title}
                        </td>
                        <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                          Found
                        </td>
                        <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                          {lesson.transcript ? "Found" : "Missing"}
                        </td>
                        <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                          {lesson.summary ? "Found" : "Missing"}
                        </td>
                        <td style={{ padding: "8px", borderTop: "1px solid #ddd" }}>
                          <button
                            onClick={() => selectLesson(lesson)}
                            style={{
                              background: selected ? "#e9e5fb" : "#fff",
                              borderColor: selected ? "#7664c5" : "#c9cbd6",
                            }}
                          >
                            {selected ? "Selected" : "Select lesson"}
                          </button>
                          {lesson.transcript && !lesson.summary && (
                            <button
                              onClick={() => generateMissingSummary(lesson)}
                              disabled={loadingSummary}
                              style={{ display: "block", marginTop: "6px" }}
                            >
                              {loadingSummary ? "Generating..." : "Generate missing summary"}
                            </button>
                          )}
                          {!complete && (
                            <div style={{ color: "#98621a", fontSize: "12px", marginTop: "4px" }}>
                              Incomplete file set
                            </div>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
            <p style={{ color: "#666", fontSize: "13px" }}>
              Matching uses the recording filename without its .mp4 extension.
              The transcript is expected to have the same filename with .txt,
              or the suffix _transcript.txt; the summary uses _summary.txt.
            </p>
          </>
        )}
      </div>
    ) : null,

    links: sharingLinks.length > 0 ? (
      <div>
        {renderPanelHeader("links", `Generated sharing links (${sharingLinks.length})`)}
        {expandedPanels.has("links") && (
          <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left" }}>
            <thead>
              <tr>
                <th style={{ padding: "8px" }}>File</th>
                <th style={{ padding: "8px" }}>Sharing link</th>
                <th style={{ padding: "8px" }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {sharingLinks.map((link) => (
                <tr key={link.id}>
                  <td style={{ padding: "8px" }}>{link.name}</td>
                  <td style={{ padding: "8px" }}>
                    <a href={link.url} target="_blank" rel="noopener noreferrer">
                      Open link
                    </a>
                  </td>
                  <td style={{ padding: "8px" }}>
                    <button onClick={() => navigator.clipboard.writeText(link.url)}>
                      Copy link
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    ) : null,

    summary: summaryText !== "" ? (
      <div>
        {renderPanelHeader("summary", "Lesson Summary")}
        {expandedPanels.has("summary") && (
          <div
            style={{
              marginTop: "10px",
              padding: "20px",
              border: "1px solid #ccc",
              borderRadius: "6px",
              textAlign: "left",
            }}
          >
            <div className="lesson-markdown">
              <ReactMarkdown
                remarkPlugins={[remarkMath]}
                rehypePlugins={[rehypeKatex]}
                components={{
                  a: ({ ...props }) => (
                    <a {...props} target="_blank" rel="noopener noreferrer" />
                  ),
                  table: ({ ...props }) => (
                    <div className="lesson-markdown-table-wrap">
                      <table {...props} />
                    </div>
                  ),
                }}
              >
                {summaryText}
              </ReactMarkdown>
            </div>
          </div>
        )}
      </div>
    ) : null,
  };

  // Restore the last OneDrive folder once per signed-in Microsoft account.
  // Only folder IDs and names are stored; Microsoft access tokens are never persisted.
  useEffect(() => {
    if (!isAuthenticated) {
      restoredFolderAccountRef.current = "";
      return;
    }

    const account = instance.getActiveAccount() || accounts[0];if (!account) return;

    const accountIdentifier = account.homeAccountId || account.username;
    const storageKey = `tutoringAutomation.lastOneDriveFolder.${accountIdentifier}`;
    if (restoredFolderAccountRef.current === accountIdentifier) return;
    restoredFolderAccountRef.current = accountIdentifier;

    const restoreLastFolder = async () => {
      let savedStack: { id: string; name: string }[];
      try {
        const saved = localStorage.getItem(storageKey);
        if (!saved) return;
        const parsed = JSON.parse(saved);
        if (!Array.isArray(parsed.folderStack) || !parsed.folderStack.every(
          (folder: unknown) => typeof folder === "object" && folder !== null &&
            typeof (folder as { id?: unknown }).id === "string" &&
            typeof (folder as { name?: unknown }).name === "string"
        )) {
          localStorage.removeItem(storageKey);
          return;
        }
        savedStack = parsed.folderStack;
      } catch (error) {
        console.warn("Unable to read the saved OneDrive folder location.", error);
        localStorage.removeItem(storageKey);
        return;
      }

      setMessage("Restoring your last OneDrive folder...");
      setFolderStack(savedStack);
      const lastFolderId = savedStack.length > 0
        ? savedStack[savedStack.length - 1].id
        : undefined;
      const restored = await loadFolder(lastFolderId);

      if (restored) {
        setFolderStack(savedStack);
        setMessage(`Restored your last OneDrive folder${savedStack.length ? `: ${savedStack.map((folder) => folder.name).join(" / ")}` : " (root)"}.`);
      } else {
        setFolderStack([]);
        localStorage.removeItem(storageKey);
        const rootLoaded = await loadFolder();
        setMessage(rootLoaded
          ? "Your previous OneDrive folder is no longer available. OneDrive root has been opened instead."
          : "Your previous OneDrive folder could not be restored. Use Load OneDrive to try again.");
      }
    };

    void restoreLastFolder();
  }, [isAuthenticated, accounts, instance]);

  return (
    <div
      style={{
        maxWidth: "1000px",
        margin: "32px auto",
        padding: "28px",
        fontFamily: "Arial, sans-serif",
        color: "#303044",
        backgroundColor: "#ffffff",
        borderRadius: "16px",
        boxShadow: "0 5px 24px rgba(35, 31, 65, 0.08)",
      }}
    >
      <style>{`
        * {
          box-sizing: border-box;
        }
        body {
          margin: 0;
          padding: 1px 18px;
          background: #f4f5fa;
          font-size: 13px;
          line-height: 1.4;
        }
        button {
          max-width: 100%;
          white-space: normal;
          overflow-wrap: anywhere;
          font-size: 12px;
          line-height: 1.35;
        }
        button {
          border: 1px solid #c9cbd6;
          border-radius: 6px;
          padding: 7px 12px;
          background: #ffffff;
          color: #303044;
          cursor: pointer;
          transition: background-color 140ms ease, border-color 140ms ease;
        }
        button:hover:not(:disabled) {
          background: #eeecf8;
          border-color: #8d80c7;
        }
        button:disabled {
          opacity: 0.55;
          cursor: not-allowed;
        }
        table {
          width: 100% !important;
          max-width: 100%;
          margin-top: 8px;
          border-collapse: collapse;
          table-layout: fixed;
          font-size: 12px;
          border-radius: 8px;
        }
        th {
          background: #f0eef8;
          color: #514879;
          font-weight: 700;
        }
        td, th {
          min-width: 0;
          padding: 6px 5px !important;
          overflow-wrap: anywhere;
          word-break: break-word;
          vertical-align: top;
        }
        td button, th button {
          max-width: 100%;
        }
        input[type="checkbox"] {
          max-width: 100%;
        }
        tbody tr:nth-child(even) {
          background: #faf9fd;
        }
        tbody tr:hover {
          background: #f1effa;
        }
        a {
          color: #4d43a5;
          font-weight: 600;
        }
        a:hover {
          color: #302675;
        }
        .status-panel {
          margin: 12px 0 18px;
          padding: 10px 16px;
          border: 1px solid #e1deee;
          border-left: 4px solid #8171c6;
          border-radius: 8px;
          background: #faf9fd;
        }
        .status-panel h3 {
          margin: 2px 0 6px;
          color: #514879;
          font-size: 16px;
        }
        .status-panel p {
          margin: 4px 0;
          color: #55536a;
        }
        .status-diagnostics-row {
          display: grid;
          grid-template-columns: repeat(3, minmax(0, 1fr));
          gap: 12px;
          align-items: stretch;
          margin: 12px 0 18px;
        }
        .status-diagnostics-row .status-panel,
        .status-diagnostics-row .diagnostics-panel { margin: 0; min-width: 0; }
        .diagnostics-panel {
          min-width: 0;
          padding: 10px 12px;
          border: 1px solid #e1deee;
          border-radius: 8px;
          background: #faf9fd;
        }
        .diagnostics-heading { display:flex; align-items:center; justify-content:space-between; gap:8px; margin-bottom:8px; }
        .diagnostics-heading h3 { margin:2px 0; color:#514879; font-size:16px; }
        .diagnostics-heading button { padding:5px 9px; font-size:12px; }
        .credit-check-list { display:grid; grid-template-columns:1fr; gap:8px; }
        .credit-check-link {
          display:block;
          width:100%;
          box-sizing:border-box;
          padding:9px 10px;
          border:1px solid #d7d2e8;
          border-radius:7px;
          background:#fff;
          color:#29235f;
          font-size:13px;
          font-weight:600;
          text-align:left;
          text-decoration:none;
          cursor:pointer;
        }
        .credit-check-link:hover { background:#f1effa; border-color:#9b90ce; }
        .credit-check-note { margin:9px 0 0; color:#625f70; font-size:11px; line-height:1.4; }
        .diagnostics-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:7px; }
        .diagnostic-item { min-width:0; padding:6px 8px; border:1px solid #e1deee; border-radius:7px; background:#fff; }
        .diagnostic-item button { display:block; width:100%; padding:2px 0; border:0; background:transparent; color:#29235f; text-align:left; font-size:13px; font-weight:600; cursor:pointer; }
        .diagnostic-item button span { margin-right:4px; }
        .diagnostic-item small { display:block; margin-top:3px; color:#625f70; font-size:11px; line-height:1.35; overflow-wrap:anywhere; }
        .diagnostics-minimized .diagnostic-item { padding:4px 7px; }
        .diagnostics-minimized .diagnostic-item small { display:none; }
        .diagnostic-ok button span { color:#16803b; }
        .diagnostic-warning button span { color:#b7791f; }
        .diagnostic-error button span { color:#c53030; }
        .diagnostic-checking button span { color:#6252b4; }
        @media (max-width: 900px) {
          .status-diagnostics-row { grid-template-columns:1fr; }
          .diagnostics-grid { grid-template-columns:1fr; }
        }
        .workflow-columns {
          display: grid;
          width: 100%;
          max-width: 100%;
          grid-template-columns: minmax(0, 1.2fr) minmax(0, 0.8fr);
          gap: 12px;
          align-items: start;
          margin-top: 14px;
        }
        .workflow-column {
          width: 100%;
          min-width: 0;
          max-width: 100%;
          overflow: hidden;
          border: 1px solid #e4e0ef;
          border-radius: 12px;
          padding: 9px;
          background: #fcfbff;
        }
        .workflow-column > div,
        .workflow-column table,
        .workflow-column p,
        .workflow-column span {
          max-width: 100%;
          overflow-wrap: anywhere;
        }
        .workflow-column-title {
          margin: 4px 4px 10px;
          color: #30286d;
          font-size: 17px;
          line-height: 1.25;
          text-align: center;
          overflow-wrap: anywhere;
        }
        .workflow-box {
          width: 100%;
          min-width: 0;
          max-width: 100%;
          margin: 8px 0 10px;
          padding: 9px;
          border: 1px solid #e3dfed;
          border-radius: 9px;
          background: #f8f7fc;
        }
        .workflow-box h3 {
          margin: 0 0 7px;
          color: #514879;
          font-size: 14px;
        }
        .workflow-status {
          display: flex;
          flex-wrap: wrap;
          gap: 5px;
          align-items: center;
          min-width: 0;
          padding: 6px 7px;
          margin: 4px 0;
          border-radius: 6px;
          font-size: 12px;
          overflow-wrap: anywhere;
        }
        .workflow-status strong {
          margin-left: auto;
          white-space: nowrap;
          font-size: 11px;
        }
        .workflow-status.done {
          color: #176534;
          background: #e4f5e9;
          border: 1px solid #b8e1c4;
        }
        .workflow-status.pending {
          color: #885400;
          background: #fff2d8;
          border: 1px solid #f0d49b;
        }
        @media (max-width: 900px) {
          .workflow-columns { grid-template-columns: minmax(0, 1fr); }
        }
      `}</style>
      <h1 style={{ marginBottom: "4px", color: "#211d3c", fontSize: "27px" }}>Tutoring Automation</h1>
      <p style={{ marginTop: "4px", color: "#77738b", fontSize: "17px" }}>
        Microsoft 365 integration
      </p>
      <hr />

      {!isAuthenticated ? (
        <button onClick={signIn}>Sign in with Microsoft</button>
      ) : (
        <>
          <p>Signed in as: {accounts[0]?.username}</p>
          <button onClick={testGraph}>Test Microsoft Graph</button>
          <button onClick={loadOneDrive} style={{ marginLeft: "10px" }}>
            Load OneDrive
          </button>
          <button
            onClick={loadOneNoteNotebooks}
            disabled={loadingNotebooks}
            style={{ marginLeft: "10px" }}
          >
            {loadingNotebooks ? "Loading notebooks..." : "Load OneNote notebooks"}
          </button>
          <button onClick={signOut} style={{ marginLeft: "10px" }}>
            Sign out
          </button>
        </>
      )}

      <hr style={{ border: 0, borderTop: "1px solid #e5e3ed", margin: "18px 0" }} />
      <div className="status-diagnostics-row">
        <div className="status-panel">
          <h3>Status messages</h3>
          <p>{message || "Ready."}</p>
        </div>
        <section className="diagnostics-panel" aria-label="Service status diagnostics">
          <div className="diagnostics-heading">
            <h3>Service status</h3>
            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <button type="button" onClick={() => setStatusMinimized((value) => !value)}>
                {statusMinimized ? "Expand status" : "Minimize status"}
              </button>
              <button type="button" onClick={runAllDiagnostics} disabled={runningAllDiagnostics}>
                {runningAllDiagnostics ? "Checking all..." : "Run all checks"}
              </button>
            </div>
          </div>
          <div className={`diagnostics-grid${statusMinimized ? " diagnostics-minimized" : ""}`}>
            {(["backend", "gemini", "assemblyai", "auth", "onedrive", "onenote", "pdf"] as DiagnosticId[]).map((id) => {
              const item = diagnostics[id];
              const icon = item.status === "ok" ? "●" : item.status === "error" ? "●" : item.status === "warning" ? "●" : item.status === "checking" ? "◌" : "○";
              return (
                <div className={`diagnostic-item diagnostic-${item.status}`} key={id}>
                  <button type="button" onClick={() => runDiagnostic(id)} disabled={item.status === "checking" || runningAllDiagnostics}>
                    <span aria-hidden="true">{icon}</span> {DIAGNOSTIC_LABELS[id]}
                  </button>
                  <small>{item.message}{item.checkedAt ? ` · ${item.checkedAt}${item.durationMs !== undefined ? ` · ${item.durationMs} ms` : ""}` : ""}</small>
                </div>
              );
            })}
          </div>
        </section>
        <section className="diagnostics-panel" aria-label="AI credit check">
          <div className="diagnostics-heading">
            <h3>AI credit check</h3>
          </div>
          <div className="credit-check-list">
            <a
              className="credit-check-link"
              href="https://aistudio.google.com/billing"
              target="_blank"
              rel="noopener noreferrer"
            >
              Google Gemini — view billing and credits ↗
            </a>
            <a
              className="credit-check-link"
              href="https://www.assemblyai.com/dashboard/settings/billing"
              target="_blank"
              rel="noopener noreferrer"
            >
              AssemblyAI — view billing and credits ↗
            </a>
            <a
              className="credit-check-link"
              href="https://portal.azure.com/#view/Microsoft_Azure_Billing/ModernBillingMenuBlade/~/Overview"
              target="_blank"
              rel="noopener noreferrer"
            >
              Azure — view credits and billing ↗
            </a>
          </div>
          <p className="credit-check-note">
            Opens the provider's billing page in a new tab. Sign in if requested; available balances depend on your account and billing permissions.
          </p>
        </section>
      </div>

      {userDetails && <pre>{userDetails}</pre>}
      {loading && <p>Loading OneDrive...</p>}
      {loadingNotebooks && <p>Loading OneNote notebooks...</p>}
      {loadingSections && <p>Loading OneNote sections...</p>}
      {loadingSummary && <p>Loading lesson summary...</p>}

      <div className="workflow-columns">
        <section className="workflow-column">
          <h2 className="workflow-column-title">OneDrive — lesson files</h2>
          {["lessons", "onedrive", "links", "summary"].map((panelId) => (
            <div key={panelId} style={{ marginTop: "10px" }}>
              {panelContent[panelId]}
            </div>
          ))}
        </section>
        <section className="workflow-column">
          <h2 className="workflow-column-title">OneNote — destination</h2>
          <div className="workflow-box">
            <h3>OneNote workflow</h3>
            {/*
              The first row is always the next incomplete action.
              Completed actions are listed underneath in reverse completion order.
            */}
            {!notebookSelected ? (
              <>
                {statusLine("1. Select a OneNote notebook", false)}
                {statusLine("2. Select a section", false)}
                {statusLine("3. Create the lesson page", false)}
              </>
            ) : !sectionSelected ? (
              <>
                {statusLine("2. Select a section", false)}
                {statusLine("1. OneNote notebook selected", true)}
                {statusLine("3. Create the lesson page", false)}
              </>
            ) : !pageCreated ? (
              <>
                {statusLine("3. Create the lesson page", false)}
                {statusLine("2. OneNote section selected", true)}
                {statusLine("1. OneNote notebook selected", true)}
              </>
            ) : (
              <>
                {statusLine("3. Lesson page created", true)}
                {statusLine("2. OneNote section selected", true)}
                {statusLine("1. OneNote notebook selected", true)}
              </>
            )}
          </div>
          {(lessonPageCreated || selectedSectionId
            ? ["createpage", "sections", "notebooks"]
            : selectedNotebookId
              ? ["sections", "notebooks"]
              : ["notebooks"]
          ).map((panelId) => (
            <div key={panelId} style={{ marginTop: "10px" }}>
              {panelContent[panelId]}
            </div>
          ))}
        </section>
      </div>
    </div>
  );
}

export default App;