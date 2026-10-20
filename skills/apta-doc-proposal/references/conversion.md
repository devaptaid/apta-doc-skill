# Conversion and troubleshooting

## Dependencies

Use Python 3.10 or newer with:

```text
markdown
beautifulsoup4
python-docx
playwright
pypdf
Pillow
reportlab
PyMuPDF
```

Install Playwright Chromium once in the active environment:

```bash
python -m playwright install chromium
```

PDF export uses Playwright for page rendering, then ReportLab and pypdf to apply the required Aptaworks running header and footer to every content page. PDF builds render a temporary PDF to map headings to pages and fill the HTML/PDF TOC. Named destinations become explicit page destinations. DOCX-only builds do not require Chromium or a temporary PDF. DOCX retains editable TOC/PAGEREF fields and same-document anchors with empty page caches, because PDF pagination is not Word pagination. It omits `w:updateFields` to avoid open-time field updates. Users refresh fields manually in Word with Ctrl+A, F9. Do not claim page numbers have been verified before that refresh.

For optional visual PDF QA, render pages with PyMuPDF (`fitz`) only after the user requests visual inspection or confirms a layout investigation. Poppler can report missing `Symbol` or `ArialUnicode` display fonts and omit or clip a valid overlay, so do not treat a Poppler-only rendering defect as proof that the PDF header is missing. When visual DOCX QA is requested or confirmed, use LibreOffice or Microsoft Word through the document renderer.

## Commands

HTML only:

```bash
python scripts/build_proposal.py proposal.md --output-dir output
```

For review-first delivery, stop here and wait for approval of Markdown/HTML. After review, convert the actual HTML without recompiling Markdown:

```bash
python scripts/build_proposal.py output/proposal.html --output-dir final --docx --approved
# Add --pdf when requested; this keeps the two-pass TOC mapping for PDF.
```

Direct final-file delivery uses the same Markdown → HTML → DOCX/PDF pipeline, without pausing between stages. Keep Markdown and HTML as editable sources. A choice of review-first overrides earlier permission to produce the eventual binary until the user asks to proceed.

Approved full export:

```bash
python scripts/build_proposal.py proposal.md --output-dir output --pdf --docx --approved
```

The binary exporters fail closed when `--approved` is missing. This is intentional.

## Common failures

- Missing package: install it in the active environment; the scripts never mutate the environment.
- Missing Chromium: run `python -m playwright install chromium`.
- Broken local image: confirm the path is relative to the Markdown/HTML file and URL-decode `%20` when checking it.
- DOCX locked: close the file in Microsoft Word, then rerun the exporter.
- Web fonts unavailable: the template falls back to Georgia, Arial, and a system monospace.
