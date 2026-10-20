#!/usr/bin/env python3
"""Run structural validation for Aptaworks HTML, PDF, and DOCX outputs."""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import unquote, urlparse


def validate_html(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    errors = []
    for marker in (
        "cover-page",
        "content-page",
        "document-header",
        "document-footer",
        "toc-page",
        "toc-entry",
        "toc-page-number",
        "logo_aptaworks_wide.svg",
        "#17234E",
        "#EF7636",
        "APTA",
        "Aptaworks",
    ):
        if marker not in text:
            errors.append(f"HTML missing marker: {marker}")
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(text, "html.parser")
        visible_text = soup.get_text(" ")
        for image in soup.find_all("img"):
            src = image.get("src", "")
            parsed = urlparse(src)
            if parsed.scheme in {"https", "http", "data"}:
                continue
            if parsed.scheme == "file":
                errors.append("HTML contains a non-portable file:// image reference.")
            elif not (path.parent / unquote(src)).is_file():
                errors.append(f"HTML image not found: {src}")
        ids = {node.get("id") for node in soup.select("[id]")}
        for link in soup.select('.toc-entry a[href^="#"]'):
            if link["href"][1:] not in ids:
                errors.append(f"HTML TOC target missing: {link['href']}")
    except ImportError:
        visible_text = re.sub(r"<[^>]+>", " ", text)
    if "AptaWorks" in visible_text:
        errors.append("HTML contains invalid branding 'AptaWorks'.")
    return errors


def validate_pdf(path: Path) -> list[str]:
    try:
        from pypdf import PdfReader
    except ImportError:
        return ["Missing dependency 'pypdf'."]
    reader = PdfReader(path)
    errors = []
    if not reader.pages:
        errors.append("PDF has no pages.")
    else:
        toc_text = reader.pages[1].extract_text() if len(reader.pages) > 1 else ""
        if not any(title in toc_text for title in ('DAFTAR ISI', 'CONTENTS')):
            errors.append("PDF is missing the contents page.")
        elif not re.search(r"(?:DAFTAR ISI|CONTENTS).*\b\d+\b", toc_text, re.DOTALL):
            errors.append("PDF TOC is missing cached page numbers.")
        annotations = reader.pages[1].get("/Annots", []) if len(reader.pages) > 1 else []
        if not annotations:
            errors.append("PDF TOC has no internal link annotations.")
        for annotation_ref in annotations:
            destination = annotation_ref.get_object().get("/Dest")
            if isinstance(destination, str):
                errors.append("PDF TOC contains an unresolved named destination.")
                break
        for page_number, page in enumerate(reader.pages[1:], start=2):
            extracted = page.extract_text() or ""
            if "Aptaworks Proposal Series" not in extracted:
                errors.append(
                    f"PDF page {page_number} does not contain the required running header."
                )
    if path.stat().st_size < 10_000:
        errors.append("PDF is unexpectedly small.")
    return errors


def validate_docx(path: Path) -> list[str]:
    errors = []
    if not zipfile.is_zipfile(path):
        return ["DOCX is not a valid ZIP/OOXML package."]
    with zipfile.ZipFile(path) as archive:
        required = {"[Content_Types].xml", "word/document.xml", "word/styles.xml"}
        missing = required - set(archive.namelist())
        if missing:
            errors.append("DOCX missing parts: " + ", ".join(sorted(missing)))
        document_xml = archive.read("word/document.xml").decode("utf-8", errors="replace")
        if "AptaWorks" in document_xml:
            errors.append("DOCX contains invalid branding 'AptaWorks'.")
        if "Aptaworks" not in document_xml:
            errors.append("DOCX does not contain Aptaworks branding.")
        if "PREPARED FOR" not in document_xml or ">APTA<" not in document_xml:
            errors.append("DOCX cover is missing the required client block or APTA watermark.")
        if 'TOC \\o "1-3"' not in document_xml or "PAGEREF" not in document_xml:
            errors.append("DOCX does not contain the editable 1-3 level Table of Contents field.")
        if 'w:hyperlink' not in document_xml or 'w:anchor="_TocApta' not in document_xml:
            errors.append("DOCX TOC does not contain same-document internal hyperlinks.")
        settings_xml = archive.read("word/settings.xml").decode("utf-8", errors="replace")
        if "w:updateFields" in settings_xml:
            errors.append("DOCX enables automatic field updates and may trigger Word's field-security prompt.")
        if not any(title in document_xml for title in ('DAFTAR ISI', 'CONTENTS')):
            errors.append("DOCX is missing the contents heading.")
        if 'w:val="Heading1"' not in document_xml:
            errors.append("DOCX is missing Heading1 heading structure.")
        # Heading 2/3 are optional; do not force invented headings into short briefs.
        root = ET.fromstring(document_xml)
        word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        for hyperlink in root.findall(f".//{{{word_ns}}}hyperlink"):
            anchor = hyperlink.get(f"{{{word_ns}}}anchor", "")
            if not anchor.startswith("_TocApta"):
                continue
            if f"{{{rel_ns}}}id" in hyperlink.attrib:
                errors.append("DOCX TOC hyperlink contains an external relationship id.")
            if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", anchor):
                errors.append("DOCX TOC hyperlink is not a valid same-document bookmark anchor.")
                break
        sect_prs = root.findall(f".//{{{word_ns}}}sectPr")
        if sect_prs:
            cover_section = sect_prs[0]
            margins = cover_section.find(f"{{{word_ns}}}pgMar")
            zero_edges = all(
                margins is not None and margins.get(f"{{{word_ns}}}{edge}") == "0"
                for edge in ("top", "right", "bottom", "left")
            )
            if not zero_edges:
                errors.append("DOCX cover section is not full bleed.")
            if cover_section.find(f"{{{word_ns}}}headerReference") is not None:
                errors.append("DOCX cover must not contain a running header.")
        header_parts = [name for name in archive.namelist() if name.startswith("word/header") and name.endswith(".xml")]
        header_xml = "".join(archive.read(name).decode("utf-8", errors="replace") for name in header_parts)
        if not header_parts or ("Aptaworks" not in header_xml and "w:drawing" not in header_xml):
            errors.append("DOCX content pages do not contain the required Aptaworks header.")
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        if not media:
            errors.append("DOCX does not contain embedded logo/image media.")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("files", nargs="+", type=Path)
    args = parser.parse_args()
    failures = 0
    validators = {".html": validate_html, ".pdf": validate_pdf, ".docx": validate_docx}
    for path in args.files:
        if not path.exists():
            print(f"FAIL {path}: file not found")
            failures += 1
            continue
        validator = validators.get(path.suffix.lower())
        if validator is None:
            print(f"FAIL {path}: unsupported file type")
            failures += 1
            continue
        errors = validator(path)
        if errors:
            failures += 1
            print(f"FAIL {path}: {'; '.join(errors)}")
        else:
            print(f"PASS {path}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
