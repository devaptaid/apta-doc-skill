---
name: apta-doc-proposal
description: Create Aptaworks A4 proposals and business documents as Markdown, branded HTML, PDF, or editable Word DOCX. Use for document deliverables; use apta-pptx for slide presentations.
---

# Aptaworks Document Proposal Builder

Create proposal content in Markdown first, compile it to HTML, then optionally export the approved HTML to PDF and DOCX.

This is a standard `SKILL.md` package for Claude Code, Codex, and compatible agents. Resolve scripts/assets from the skill installation directory rather than relying on agent-specific tools or a hardcoded working directory.

## Select the deliverable first

- If the format is missing, ask once in the user's language: “Which output format would you like: Markdown (.md), PDF, Word (.docx), or PowerPoint (.pptx)? You can choose more than one.” Wait for the format choice before exporting. You may draft the content while waiting.
- Offer the delivery mode at the initial brief: **finish the Word/PDF now**, or **prepare Markdown + HTML for review first**. Combine this with the format question if needed. Honor an already explicit mode without repeating the question. Direct delivery still runs Markdown → HTML → DOCX/PDF and keeps both sources; it skips the review pause. Review-first stops after Markdown/HTML, even if Word/PDF is the eventual target, until the user asks to proceed.
- Honor a format already specified in the current request or established conversation; do not ask again. A request for Word, PDF, or all document formats authorizes those exports and supplies the `--approved` gate.
- Markdown only: write the source and stop. Do not compile HTML or run binary exporters.
- HTML: compile and validate HTML. PDF/Word: keep Markdown and HTML as sources, then export only the requested binaries.
- After HTML review, pass that file to `build_proposal.py reviewed.html --output-dir output --docx --approved` (add `--pdf` when requested). It exports the reviewed HTML without rereading Markdown and retains the two-pass PDF TOC process. Do not recompile Markdown over approved HTML edits. Keep later source revisions synchronized deliberately.
- PowerPoint: use `apta-pptx` when installed. If unavailable, explain how to install it from the repository README; do not pass slide content through the A4 compiler. Each skill installs independently.
- For both Word and slides, make separate document and slide sources from the same approved facts. Do not squeeze A4 pages into slides. Word means `.docx`, not legacy `.doc`.
- Invoke scripts using absolute paths resolved from this skill's installed directory. The user's working directory is the output project, not the skill installation. Commands below assume you are in the skill directory.

## Required workflow

1. Read the user's brief and create or update the Markdown source.
2. Use Aptaworks casing exactly. Never write `AptaWorks` or identify the author as `Aptaworks Consultant`; use `Aptaworks Team`.
3. Treat the Markdown headings as the numbering source. Write semantic headings without manual section numbers; the compiler adds the visible hierarchy automatically:
   - `# Heading` becomes `1. Heading`.
   - `## Subheading` becomes `1.1. Subheading`.
   - `### Detail` becomes `1.1.1. Detail`.
   Existing leading numbers are stripped before numbering so headings never become `1. 1. Heading`.
4. For formatted documents, include a dedicated contents page immediately after the cover and before the first content heading: `CONTENTS` for `language: en`, `DAFTAR ISI` for `language: id`. List existing Heading 1 through Heading 3 using numbered display text, dotted leaders, and links/bookmarks. HTML/PDF cache page numbers after PDF pagination. DOCX uses editable `TOC \o "1-3" \h \z \u` and `PAGEREF` fields with explicit internal `w:hyperlink` anchors; refresh in Word for its own page numbers. Do not invent subheadings just to populate three levels.
5. When HTML, PDF, or DOCX was selected, compile Markdown to HTML without asking for permission again. Skip this step for Markdown-only:

   ```bash
   python scripts/build_proposal.py path/to/proposal.md --output-dir path/to/output
   ```

6. Validate HTML structure, local asset resolution, and branding before offering binary exports. Perform browser/screenshot inspection only when the user requests visual QA or after asking for confirmation because a suspected layout issue cannot be verified structurally. If confirmed and an interactive browser blocks `file://` URLs, run `python scripts/render_html_preview.py path/to/proposal.html --output path/to/preview.png`. A preview screenshot does not grant PDF/DOCX permission.
7. Treat PDF and DOCX as gated exports:
   - Proceed when the user explicitly asks for PDF, Word, DOCX, or all formats.
   - Otherwise ask for permission before running either export.
   - Never infer permission merely because HTML was requested.
8. After permission, export with the deterministic approval flag:

   ```bash
   python scripts/build_proposal.py path/to/proposal.md --output-dir path/to/output --pdf --docx --approved
   ```

9. Validate all requested artifacts:

   ```bash
   python scripts/validate_outputs.py path/to/output/proposal.html path/to/output/proposal.pdf path/to/output/proposal.docx
   ```

10. Validate all requested artifacts structurally by default. Do not render PDF/DOCX pages or call image-view tools as a mandatory step. Render and inspect every page only when the user explicitly requests visual QA or confirms a layout investigation; then check the TOC page, numbering, page references, clipping, tables, glyphs, and header/footer placement.

## Source contract

Start proposal Markdown with simple YAML frontmatter:

```markdown
---
title: Proposal Title
language: en
subtitle: Concise value proposition
client: Client Name
client_logo: assets/client-logo.png
author: Aptaworks Team
date: 7 August 2026
version: 1.0
status: Confidential
display_number: 01
---
```

Write semantic Markdown beneath the frontmatter. Use headings, real lists, blockquotes for callouts, fenced code blocks, images with relative paths, and tables only for comparable row-and-column data.

The bundled example is English. Set `language: en` or `language: id` to match the requested document language; this localizes the contents title, cover metadata labels, and PDF pagination. Omitted language keeps the original Indonesian furniture for backward compatibility. Translate the source content as well as selecting its language.

Do not manually number Markdown headings. Use `#`, `##`, and `###` to define the hierarchy; numbering is generated consistently in HTML, PDF, and Word. Keep the first three heading levels meaningful because all three appear in the Table of Contents.

## Resource routing

- Run `scripts/build_proposal.py` for the normal end-to-end workflow.
- Run `scripts/md_to_html.py` when only HTML is needed.
- Run `scripts/render_html_preview.py` when local browser security prevents direct HTML inspection.
- Run `scripts/html_to_pdf.py` or `scripts/html_to_docx.py` only after the export permission gate is satisfied.
- Read `references/design-system.md` when changing visual tokens or page geometry.
- Read `references/conversion.md` when installing dependencies or diagnosing conversion failures.
- Use `assets/proposal-template.html` and `examples/example.html` as the portable HTML/PDF visual authority. When a project provides `proposal/proposal-simple.html`, use it as the upstream reference.
- Use `examples/example.docx` as the portable Word cover and section-geometry authority. When a project provides `proposal/proposal-simple.docx`, use it as the upstream reference. Reproduce its full-bleed cover with paragraph indents rather than relying on Word table-cell margins.
- Use only the official wide wordmark assets: `assets/logo_aptaworks_wide.svg`, `assets/logo_aptaworks_wide_navy.png`, and `assets/logo_aptaworks_wide_white.png`. Do not redraw, approximate, or replace the Aptaworks logo.
- Use `examples/example.md`, `examples/example.html`, `examples/example.pdf`, and `examples/example.docx` as a known-good output set.
- When a project provides `proposal/proposal-simple.docx`, use its `DAFTAR ISI` page as the TOC authority: Heading 1/2/3 entries, dotted leaders, numbered labels such as `1.`, `1.1.`, and `1.1.1.`, and editable Word page references.

## Quality gates

- Keep A4 portrait geometry: 210 × 297 mm, 25 mm left/right margins, and 30 mm top/bottom margins on content pages.
- Keep the cover full bleed and use the navy, blue, and orange Aptaworks palette from the template.
- In DOCX, align the cover from the top like `examples/example.docx`: wide logo, proposal kicker, title, subtitle, `PREPARED FOR` client mark, large `APTA` watermark, and right-aligned author/date/version metadata.
- Show the official wide Aptaworks wordmark on the cover and in the content-page header.
- Keep the cover free of running furniture. Show a navy/orange running header and restrained footer on every content page in PDF and DOCX. Show an equivalent visible header/footer around the HTML content.
- Match the content typography, spacing, tables, blockquotes, images, and heading scale in `assets/proposal-template.html`; do not substitute a generic proposal layout.
- Preserve heading hierarchy and editable Word structure; do not create a DOCX by embedding page screenshots.
- Always generate a TOC page after the cover. Use Word Heading 1/2/3 styles with outline levels 0/1/2 and keep the TOC field limited to levels 1–3. The TOC title must not itself use a Heading 1–3 style, otherwise it will appear in its own contents list.
- For PDF, use a two-pass export to map heading pages and fill the HTML/PDF TOC. Repair named destinations to explicit page destinations. DOCX has different pagination: never copy PDF page numbers into Word fields. Keep editable `TOC`/`PAGEREF` fields with same-document anchors and an empty page cache until refreshed in Word. Omit `w:updateFields`; explain Ctrl+A, F9 in the handoff. Only claim verified Word page numbers after a Word-native field refresh and pagination check.
- Never leave both a manually numbered heading and a compiler-generated number in the output.
- Resolve relative images from the Markdown or HTML file directory, including URL-encoded spaces.
- Keep source Markdown and HTML as the editable source of truth.
- Do not install packages automatically from the conversion scripts. Report missing dependencies and let the user or environment owner decide how to install them.
