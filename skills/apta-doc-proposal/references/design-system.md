# Aptaworks A4 Editorial Design System

Use bundled `assets/proposal-template.html` and `examples/example.html` as the portable HTML/PDF visual authority, and `examples/example.docx` as the portable DOCX cover and section-geometry authority. When the active project provides `proposal/proposal-simple.html` or `proposal/proposal-simple.docx`, treat those files as upstream references. Use these tokens consistently in HTML, PDF, and DOCX.

| Role | Value |
|---|---|
| Navy | `#17234E` |
| Deep navy | `#0B1533` |
| Orange accent | `#EF7636` |
| Blue accent | `#0A6F93` |
| Warm callout | `#F7F5EE` |
| Rule | `#C9D2DF` |
| Muted text | `#5E6B82` |

Use Newsreader or Georgia for display headings, Lato or Arial for body text, and DM Mono or a system monospace for code. Match the reference scale: 12.5 pt body, 26 pt H1, 20 pt H2, and 14 pt H3. Keep the document readable when web fonts are unavailable.

Use an A4 portrait page. Keep the cover full bleed. On content pages use 25 mm horizontal and 30 mm vertical margins. Avoid splitting headings from their first paragraph, rows across pages where possible, and code/callout blocks across pages.

Use the official 314×63 wide Aptaworks wordmark. Use the white version on the navy cover and the navy version in content headers. Never construct a replacement logo from text or geometric shapes.

For the Word cover, use a full-page fixed table with zero section margins and zero cell margins. Apply 25 mm left/right paragraph indents because Word does not reliably honor cell padding on a 100%-width cover table. Lay out the wide logo from the top, then the proposal kicker, title, subtitle, prepared-for client mark, `APTA` watermark, and right-aligned metadata near the bottom.

Keep running headers off the cover. On each content page show status and document title at left, then `Aptaworks Proposal Series` plus the wide wordmark at right, with an orange rule beneath. Use a muted footer separated by a light rule.

## Table of contents and heading numbering

Place a dedicated contents page immediately after the cover: `CONTENTS` for `language: en`, `DAFTAR ISI` for `language: id`. English examples use English cover metadata and PDF pagination. Number the first three semantic heading levels automatically and use the same display text everywhere:

```text
1. Executive Summary
3. Objectives and Outcomes
3.1. Project Objectives
9.3.1. Cost Breakdown
```

The Word output defines Heading 1, Heading 2, and Heading 3 styles (outline levels 0, 1, and 2); use only the levels the content needs. The TOC must be an editable field equivalent to `TOC \o "1-3" \h \z \u`, with `TOC 1`, `TOC 2`, and `TOC 3` paragraph styles, dotted tab leaders, and `PAGEREF` page references. Wrap each visible TOC entry in a same-document `w:hyperlink` with a `w:anchor` matching the heading bookmark; do not create external TOC hyperlink relationships. Leave cached Word page numbers empty until refreshed in Word; do not reuse PDF pagination. Omit `w:updateFields`; users refresh fields manually with Ctrl+A, F9.

Do not number headings in Markdown. If a source heading already begins with a number, strip that prefix before applying the generated number.

Use Aptaworks casing exactly. Use `Aptaworks Team` for the author/team label.
