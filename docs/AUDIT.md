# Aptaworks skills audit

Updated: 20 October 2026. Scope: skills CLI / Claude Code installation, Markdown → HTML → final-file workflows, portable generation, source coverage, English examples, and clean presentation styling.

This is a local engineering audit, not a skills.sh certification. Structural checks and example rendering do not guarantee that every future brief will fit without review.

## Current package

Two independently installable skills live in `skills/apta-doc-proposal` and `skills/apta-pptx`. Each includes scripts, assets, references, and examples. The README starts with installation commands for **[devaptaid/apta-doc-skill](https://github.com/devaptaid/apta-doc-skill/)** and explicitly supports `--agent claude-code`.

The real `skills` CLI successfully installed both packages into an isolated Claude Code project using its default installation mode. Both installed generators ran successfully from `.claude/skills/`, including document HTML/DOCX and presentation HTML/PPTX output. This verifies package installation and script execution, not a full Claude conversation/behavioral evaluation.

**Remote publication blocker:** `npx skills add devaptaid/apta-doc-skill --list` cloned the repository but returned “No skills found.” GitHub reported the repository empty (`isEmpty: true`); the active GitHub identity had `viewerPermission: READ`. The local package must be published to the repository's default branch with write access before the `owner/repo` shorthand can install it. Do not represent the remote installation as successful yet.

The document example is an English, illustrative A4 proposal. The presentation is a 13-slide English component library: cover, ledger, KPI metrics, workstream cards, flow, timeline, divider, native table, two-column proposal, proof grid, source image, quote, and consultation closing.

## Findings and corrections

| Finding | Correction | Evidence |
|---|---|---|
| HTML preview navigation differed from the original slider | Restored the bottom 7px dots and 22px orange active pill, with direct selection, keyboard, wheel throttling, and swipe navigation | Chromium interaction checks, portrait viewports, print visibility, and screenshots of all 13 slides |
| Large titles and spaced header labels reduced content space | Content/section titles are 36pt, cover titles 48pt; metadata and eyebrow move upward, and components start at y=184 or y=244 according to title wrapping | Actual PowerPoint render reviewed across the 13-slide library; HTML computed typography and text bounds checked |
| Unwanted “Editorial” labels appeared in the presentation | Header brand is `Aptaworks`; removed the cover edition and generated Editorial labels | Browser check scans every generated slide; English cover copy updated |
| Presentation backgrounds were too busy | Removed the paper grid from all generated HTML/PPTX slides while retaining brand typography, gradients, and information-bearing lines | Native output regression checks every slide for background-grid geometry |
| PPTX had no generated HTML stage | Added self-contained positioned HTML and made the native exporter consume its visible DOM | Regression deletes the Markdown, edits/moves HTML, then verifies changed text/typography in PPTX |
| Review-first content could be regenerated from Markdown before export | Both pipelines accept reviewed HTML directly | Document and presentation regressions verify reviewed HTML changes survive without Markdown |
| Delivery timing was unclear | Offer direct final output or Markdown + HTML review at the initial brief | Both routes share the complete pipeline; review-first stops before binary output |
| The first closing moved the QR to a side column and omitted contact details | Centered headline/subtitle/QR; restored the navy-to-cyan gradient and original complete contact artwork | Closing preview; regression verifies gradient stops, QR centering, and identical embedded footer bytes |
| The deck had too few examples of common presentation components | Added cards, KPI metrics, timeline, proof grid, quote, and a detailed ledger alongside tables and flows | 13-slide native output; source-to-output coverage validator |
| English content still inherited Indonesian document furniture | Added `language: en` / `language: id`; localized contents, author/date labels, and PDF page labels | English DOCX/HTML regression; English PDF and Word render |
| Word could leave the last item of a short list alone on the next page | Short lists keep their items together | Updated Word example rendered and inspected |
| HTML referenced images through local file URLs | Bundled adjacent output assets and checked their resolution | Relocated-output regression |
| Word page references used PDF pagination | Word retains its own editable fields and refresh workflow | Word and PDF paginate differently; PDF page numbers are no longer reused for Word |
| The original validator required all heading levels in short documents | Heading 1 is required; Heading 2/3 follow actual content | Single-heading document regression |
| Early PPTX tables used the wrong schema URI | Corrected the native table schema URI | PowerPoint opens/exports the deck; validator checks the URI |

## Verification

| Check | Result |
|---|---|
| Skill frontmatter/naming validation | Both skills pass `quick_validate.py` |
| Local skills CLI discovery | Both skills found with `npx skills add . --list` |
| Claude Code installation | Real CLI default-mode installation into an isolated project; both installed generators pass |
| Remote shorthand discovery | Blocked: the supplied GitHub repository is empty and the active identity has READ access |
| Standalone copied installation | Document and PPTX builders work from isolated directories, without a sibling skill |
| Document generation | English Markdown → HTML, PDF, DOCX |
| Document structural validation | Assets, TOC, fields, bookmarks, geometry, headers, and embedded media pass |
| Native document rendering | Microsoft Word 16 renders the English DOCX; PDF and Word are each 4 pages |
| Presentation generation | English Markdown → self-contained HTML → 13-slide native PPTX |
| Presentation structural validation | Relationships, bounds, source text, notes, native table, and closing assets pass |
| Native presentation rendering | Microsoft PowerPoint 16 opens and exports all 13 slides |
| HTML slider behavior | Chromium verifies dots, Enter/Space activation, arrow/Home/End keys, wheel throttling, swipe, and print mode |
| Responsive preview | 1280×720, 786×1051, and 375×812; slide canvas and dot navigation stay within the viewport |
| Visual review | Every slide and every document PDF/Word page inspected; no clipping or overlap in the final examples |
| Regression suite | Thirteen tests pass, including multiple invalid-input cases |

The tests cover standalone installation, missing assets, unsupported layouts/syntax, excessive density, malformed tables, source text loss, document export approval flags, relocated HTML assets, a brief with only Heading 1, English document labels, and the reference closing artwork/geometry. The closing contact row is preserved as supplied artwork; headline text remains editable.

The source image example uses the official wordmark, with no grid in the artwork. Flows, cards, KPI cells, timelines, and proof grids are native editable PowerPoint objects. The PDF preview is exported from the actual PPTX rather than from HTML. Presentation HTML is self-contained; its embedded images, text, inline geometry, and notes are parsed without executing scripts or fetching URLs.

## Output-routing review

This is an instruction/routing review, not an independent agent forward-test.

| Request | Required behavior |
|---|---|
| Format not specified | Ask once: Markdown, PDF, Word, or PowerPoint; allow multiple selections |
| Delivery mode not established | Offer direct final output or Markdown + HTML for review; honor an explicit choice |
| Markdown + HTML for review | Compile only the review artifacts, then wait before binary export |
| Continue after HTML review | Export the reviewed HTML rather than regenerating it from Markdown |
| Markdown only | Write the source and stop before compilation |
| Word and PDF | Use the document pipeline; the explicit request satisfies the export gate |
| PPTX with a specified slide count | Preserve the total count, including cover/closing |
| PDF slides | Render the deck through PowerPoint/LibreOffice; do not substitute A4 pages |
| Word and PPTX | Separate sources/layouts using the same approved facts |
| Missing counterpart skill | Explain its installation rather than assuming a sibling folder exists |

## Remaining limits

- **Word contents:** refresh with Ctrl+A, F9 to populate Word-native page numbers, then refresh after editing. Rendering may calculate fields in the preview session; the distributed source is not silently saved with guessed PDF page numbers.
- **Fonts:** Lato was available during review. Fonts are not embedded; substitutions can change wrapping. The native implementation uses Georgia/Consolas in the display/metadata roles occupied by Newsreader/DM Mono in HTML.
- **Compiler scope:** documented Markdown and positioned-HTML subset, embedded PNG, native tables, and editable diagram primitives. Native charts, rich text, animations, SVG input, and arbitrary web-page conversion are not implemented. Base stylesheet/script edits are browser-only; supported element styling is inline.
- **Design matching:** the closing preserves the source's gradient stops, hierarchy, QR position, and original footer artwork, with the background grid removed. Browser text rasterization/font metrics can differ from PowerPoint.
- **Future output:** inspect each new deck in a slide renderer; the example does not cover every possible title length, table density, language, or Office version.
- **Distribution:** the repository is still empty remotely and the active account has READ access; publication and a successful GitHub-source installation remain outstanding. Local Claude Code installation is verified. Brand assets remain subject to their owner's rights; no general redistribution license has been assigned here.

## Reproduce

```bash
python -m unittest discover -s tests -v
python tests/check_deck_browser.py --output .build/deck-browser-check
npx skills add . --list
python tests/check_install.py --workspace .build/claude-install-check
python tests/check_install.py --source devaptaid/apta-doc-skill --workspace .build/claude-github-check
python skills/apta-pptx/scripts/validate_deck.py skills/apta-pptx/examples/example.pptx --source skills/apta-pptx/examples/example.md
python skills/apta-pptx/scripts/validate_deck.py skills/apta-pptx/examples/example.html
python skills/apta-doc-proposal/scripts/validate_outputs.py skills/apta-doc-proposal/examples/example.html skills/apta-doc-proposal/examples/example.pdf skills/apta-doc-proposal/examples/example.docx
```

For optional native Windows QA with Office and `pywin32` already installed:

```bash
python tests/render_office.py input.pptx new-output.pdf
python tests/render_office.py input.docx new-output.pdf
```

The Office helper refuses to overwrite an existing PDF and closes only the artifact it opened. It does not save changes to the source document. Browser QA is optional and requires an existing Playwright/Chromium installation; it writes HTML screenshots to the chosen output directory.
