# ClearBrief

A polished, local-first text summarizer and document analysis app built with Streamlit. It extracts source wording and ranks the most informative sentences; it does not send documents to an external AI service.

## Features

- Paste text or upload multiple documents: TXT, Markdown, RST, LOG, CSV/TSV, JSON, YAML, XML, HTML, RTF, PDF, DOCX, PPTX, XLSX, and common image formats.
- Extract text from PDFs, Word documents and tables, presentation slides, spreadsheets, web pages, and images (OCR).
- Automatically summarize the selected uploaded document using the chosen output format and length.
- Choose balanced, bullet-point, executive-brief, or action-item output in brief, standard, or detailed lengths.
- Boost sentences related to an optional focus topic.
- Live word, character, sentence, paragraph, reading-time, and speaking-time counts.
- Inspect readability estimates, frequent terms, sentence length, and summary compression.
- Summarize multiple uploaded files and download summaries as plain text, Word documents, or PDFs.
- Save useful summaries to a session-only library.

## Run locally

1. Create and activate a Python 3.10+ virtual environment.
2. Install dependencies with `pip install -r requirements.txt`.
3. Launch with `streamlit run app.py`.

Image OCR also requires the Tesseract OCR application installed on the computer. Without it, all other supported formats remain available.

## Sample text dataset

Choose one of the long-form practice passages in the app's **Sample text** menu and select **Load sample text** to try summarization. The passages cover urban tree planning, community river monitoring, and public library digital access. The texts are illustrative examples, not reports of real projects.

## Notes

- Uploads are limited to 50 MB per file.
- Processing uses the app's Python process. If hosted remotely, uploaded files are processed by that host; deploy only where your data-handling requirements are met.
- PDF text extraction works best on text-based PDFs. Scanned PDFs may need OCR before upload.
- The summarizer is extractive and works offline after setup. It selects sentences from the original, so the result is best treated as a concise reading aid, not a substitute for source review.
- Session library data is temporary and is cleared when the app session ends.