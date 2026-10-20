# Generation and visual QA

## Dependencies

Markdown → HTML → PPTX generation and structural validation need only Python 3.10+. The HTML exporter reads the visible positioned DOM and writes standard OOXML with native text/tables; no private artifact runtime, node packages, `python-pptx`, or Microsoft Office license is needed for generation. See [HTML conversion](html-conversion.md). The renderer is a separate dependency. No script installs packages automatically.

For the closest visual match, install Lato (including Lato Black and Bold) on the authoring and recipient computers. The native example uses Georgia italic for display figures and Consolas for metadata. Font files are not bundled or embedded. The packaged glyph advances support density checks without installing dependencies, but cannot prevent font substitution. Say when a required font is unavailable and inspect the rendered result.

## Rendering the actual deck

Use an installed PowerPoint or LibreOffice instance. Do not export the upstream HTML to pretend that the PPTX has been inspected.

LibreOffice example (separate fresh output folder to avoid overwrite prompts):

```bash
soffice --headless --convert-to pdf --outdir /path/to/preview /path/to/deck.pptx
pdftoppm -scale-to 1600 -png /path/to/preview/deck.pdf /path/to/preview/slide
```

On Windows, PowerPoint can open the generated deck and export PDF with its native exporter. Keep renderer automation in the build/project workspace; do not change the user's running document. Verify the exported PDF has the same slide count. Inspect all pages; if a PowerPoint repair prompt appears, treat it as failure and fix the package.

If no slide renderer is available, run `validate_deck.py` and explain that visual fit is not verified. A passing structural check is not a visual sign-off.

## Boundaries

`build_deck.py source.md --output deck.html` creates the review artifact only. `build_deck.py deck.html --output deck.pptx` exports the reviewed HTML. `build_deck.py source.md --output deck.pptx` runs both stages and retains HTML. Markdown-only does not invoke it. PDF slide export is via a slide renderer; A4 PDF/Word belongs to `apta-doc-proposal`. Arbitrary browser-page conversion, animation support, automatic chart generation, SVG rendering, and legacy `.ppt` output are not implemented.
