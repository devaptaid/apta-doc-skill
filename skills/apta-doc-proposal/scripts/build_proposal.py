#!/usr/bin/env python3
"""Build Aptaworks proposal HTML and optionally approved PDF/DOCX exports."""

from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path

from html_to_docx import export_docx, extract_heading_page_numbers
from html_to_pdf import export_pdf
from md_to_html import apply_toc_page_numbers, compile_markdown, package_image_sources


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="Markdown source or reviewed proposal HTML")
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--pdf", action="store_true")
    parser.add_argument("--docx", action="store_true")
    parser.add_argument("--approved", action="store_true", help="Confirm explicit user permission for PDF/DOCX export.")
    args = parser.parse_args()
    if (args.pdf or args.docx) and not args.approved:
        print("ERROR: PDF/DOCX export requires explicit user approval (--approved).", file=sys.stderr)
        return 2
    skill_dir = Path(__file__).resolve().parent.parent
    source = args.source.resolve()
    output_dir = args.output_dir.resolve()
    stem = source.stem
    html_path = output_dir / f"{stem}.html"
    try:
        if source.suffix.lower() == '.md':
            compile_markdown(source, html_path, skill_dir / "assets" / "proposal-template.html", skill_dir / "assets" / "logo_aptaworks_wide.svg")
        elif source.suffix.lower() == '.html':
            if source != html_path:
                output_dir.mkdir(parents=True, exist_ok=True)
                markup = package_image_sources(source.read_text(encoding='utf-8'), source.parent, output_dir)
                html_path.write_text(markup, encoding='utf-8')
        else:
            raise ValueError('Input must be .md or reviewed .html.')
        print(f"HTML: {html_path}")
        page_map: dict[str, int] = {}
        if args.pdf:
            pdf_for_toc = output_dir / f".{stem}.toc-cache.pdf"
            try:
                asyncio.run(export_pdf(html_path, pdf_for_toc, approved=True))
                page_map = extract_heading_page_numbers(pdf_for_toc, html_path)
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(html_path.read_text(encoding="utf-8"), "html.parser")
                expected = {entry.get("data-toc-display") for entry in soup.select(".toc-entry")}
                if not expected or expected - page_map.keys():
                    raise RuntimeError("Could not resolve every PDF TOC page. Check heading layout before export.")
            finally:
                if pdf_for_toc.exists():
                    pdf_for_toc.unlink()
            apply_toc_page_numbers(html_path, page_map)
        if args.pdf:
            result = asyncio.run(export_pdf(html_path, output_dir / f"{stem}.pdf", approved=True))
            print(f"PDF: {result}")
        if args.docx:
            result = export_docx(
                html_path,
                output_dir / f"{stem}.docx",
                approved=True,
                page_map={},
            )
            print(f"DOCX: {result}")
    except (OSError, ValueError, RuntimeError, PermissionError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
