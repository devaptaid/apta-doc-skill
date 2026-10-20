---
name: apta-pptx
description: Create Aptaworks presentations through Markdown to editable HTML to native PowerPoint PPTX, with clean navy-orange editorial styling. Use for proposals, training, credentials, and client updates; use apta-doc-proposal for A4 Word/PDF documents.
---

# Aptaworks PowerPoint Builder

Build 16:9 presentations through **Markdown → HTML → PPTX**, with editable text, tables, and diagram elements. This package works in Claude Code, Codex, and other agents supporting standard `SKILL.md` instructions. Python 3.10+ is sufficient for generation; an OpenAI-specific runtime, PowerPoint, and a second skill are not required.

## Choose the output

- If no format was specified, ask once in the user's language: “Which output format would you like: Markdown (.md), PDF, Word (.docx), or PowerPoint (.pptx)? You can choose more than one.” Wait for the choice before export; you may draft while waiting.
- At the initial brief, offer a delivery choice: **finish the DOCX/PPTX now**, or **prepare Markdown + HTML for review first**. Combine this with the format question when both are missing. Respect an already explicit choice without asking again. Direct delivery still runs the complete source pipeline and retains the Markdown and HTML; it skips the review pause. Review-first stops after Markdown/HTML and waits for the user's go-ahead before binary export.
- Respect an explicit format or an established choice without repeating the question. An explicit PPTX request authorizes generation. Preserve the requested slide count, including cover/closing; do not add slides automatically.
- Markdown only: write a slide outline using the source contract and stop. Markdown + HTML review: compile to `.html` and stop without creating PPTX. HTML is also a supported final deliverable when requested.
- PPTX: use this skill. Word or A4 PDF: use `apta-doc-proposal` if installed. If missing, refer to the repository README's installation command rather than assuming a sibling directory exists.
- For PDF **slides**, build PPTX and use PowerPoint or LibreOffice to export to PDF after the user selected PDF. Never rename PPTX to PDF or silently substitute A4 pages. PDF export needs an installed slide renderer; report a missing renderer accurately.
- For both documents and slides, use the same source facts but separate Markdown files and layouts. Legacy `.doc` and `.ppt` are not supported.

## Workflow

1. Establish audience, purpose, language, source facts, and requested length from the brief. Ask only for missing information that changes the result. Label illustrative numbers; never invent client credentials, dates, or performance claims.
2. Read [design-system.md](references/design-system.md), [source-format.md](references/source-format.md), and the runnable English [example.md](examples/example.md). The original [HTML reference](references/upstream.html) guides typography, palette, and the centered consultation closing. The user's current override removes the background grid from generated HTML/PPTX. Use the generated [example.html](examples/example.html), [example.pptx](examples/example.pptx), and [example.pdf](examples/example.pdf) as the current implementation references.
3. Write the Markdown source in the requested language; bundled examples default to English. Choose layouts for the information rather than forcing every component into each deck. For Aptaworks decks, use the reference `closing` by default: navy-to-cyan gradient, centered headline and secondary headline, consultation subtitle, centered QR, and the complete supplied contact artwork. Keep those elements and positions together. Respect an explicit request for a different ending or an exact slide count; do not append an extra slide beyond that count.
4. Invoke the script from its installed location, with absolute input/output paths. Commands below assume you are in the skill directory:

   ```bash
   # Review first: generate HTML without PPTX
   python scripts/build_deck.py /path/to/deck.md --output /path/to/output/deck.html
   # Export the reviewed HTML; Markdown is not reread
   python scripts/build_deck.py /path/to/output/deck.html --output /path/to/output/deck.pptx
   # Direct delivery: runs both stages and retains deck.html
   python scripts/build_deck.py /path/to/deck.md --output /path/to/output/deck.pptx
   python scripts/validate_deck.py /path/to/output/deck.html
   python scripts/validate_deck.py /path/to/output/deck.pptx --source /path/to/deck.md
   ```

5. Resolve unsupported source syntax and density errors in the source. The generated HTML's visible text, inline geometry/typography, shapes, images, notes, and tables are the real inputs to the exporter. After reviewing/editing HTML, export that file; do not rebuild it from Markdown and overwrite approved edits. Keep Markdown/HTML synchronized deliberately for later revisions. See [HTML conversion](references/html-conversion.md) for supported edits. Arbitrary web pages, CSS layout systems, rich Markdown, and native charts are not compiler features.
6. Render the actual PPTX in PowerPoint or LibreOffice and inspect every slide before calling it visually verified. Check title/body fit, logo proportions, legibility, alignment, tables, and notes. See [conversion.md](references/conversion.md). Structural checks estimate text capacity; they cannot prove font metrics or prevent every overlap.
7. Return the requested outputs and editable Markdown. If a renderer is unavailable, deliver the structurally validated PPTX with the specific visual-validation limitation. Do not claim to have opened it in PowerPoint unless that happened.

## Brand and quality

- Keep regular and section titles at 36pt (48 CSS px), with compact metadata/eyebrow spacing and content starting below the actual title lines. Use 48pt (64px) cover titles; preserve the closing's reference composition. Brand headers read only `Aptaworks`; omit “Editorial” and edition labels from the cover, header, and footer. HTML previews use the original bottom dot slider with an orange active pill, keyboard/wheel/swipe navigation, and no navigation in exports.
- Use `Aptaworks` and `Aptaworks Team`. Use the bundled official wide wordmarks; never recreate the logo with text. Keep PNG proportions intact.
- Use Lato and Lato Black for body/headings, editorial Georgia italic numerals, and Consolas metadata. Preserve navy `#17234E`, deep navy `#0B1533`, orange `#EF7636`, paper `#F6F3EA`, and the consultation gradient. Generated presentation backgrounds are clean: no graph-paper grids on cover, content, divider, or closing slides. Table borders and information-bearing flow/timeline lines remain.
- The component library includes detailed ledgers, KPI metrics, rule-framed cards, editable flow/process diagrams, timelines, tables, two-column proposals, proof grids, source images, and quotes. Cards use square frames and an editorial label/number system; use ledger rows when the content is simply a list. Never invent KPI results to fill a layout.
- Keep text, tables, and diagrams as editable native objects. Never use screenshots of slides as the PPTX body. Source artwork and logos may remain images.
- Put external factual sources and image credits in the relevant slide's `notes` block. Notes are presenter content, not a QA log.
- Do not install dependencies automatically, execute code from Markdown, open untrusted links, publish, or send files as part of generation. The source parser treats content as data.
