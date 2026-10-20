#!/usr/bin/env python3
"""Export Aptaworks proposal HTML to PDF after explicit approval."""

from __future__ import annotations

import argparse
import asyncio
import io
import re
import sys
from pathlib import Path


def _repair_internal_links(pdf_path: Path, html_path: Path) -> None:
    """Replace Chromium named destinations with explicit page destinations."""
    from bs4 import BeautifulSoup
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import ArrayObject, NameObject, NullObject

    soup = BeautifulSoup(html_path.read_text(encoding="utf-8"), "html.parser")
    toc_items = {
        entry.get("data-toc-id", ""): entry.get("data-toc-display", "")
        for entry in soup.select(".toc-entry")
        if entry.get("data-toc-id")
    }
    if not toc_items:
        return

    reader = PdfReader(str(pdf_path))
    page_texts = [re.sub(r"\s+", " ", page.extract_text() or "").strip().lower() for page in reader.pages]
    first_display = next(iter(toc_items.values()), "").lower()
    occurrences = [index for index, text in enumerate(page_texts) if first_display in text]
    body_start = occurrences[1] if len(occurrences) > 1 else (occurrences[0] + 1 if occurrences else 1)
    page_by_id: dict[str, int] = {}
    for section_id, display in toc_items.items():
        target = re.sub(r"\s+", " ", display).strip().lower()
        for page_index in range(body_start, len(page_texts)):
            if target in page_texts[page_index]:
                page_by_id[section_id] = page_index
                break
    if not page_by_id:
        return

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    changed = False
    for page in writer.pages:
        annotations = page.get("/Annots")
        if not annotations:
            continue
        for annotation_ref in annotations:
            annotation = annotation_ref.get_object()
            destination = annotation.get("/Dest")
            if not isinstance(destination, str):
                continue
            target_index = page_by_id.get(destination.lstrip("/"))
            if target_index is None:
                continue
            target_page = writer.pages[target_index]
            annotation[NameObject("/Dest")] = ArrayObject(
                [target_page.indirect_reference, NameObject("/XYZ"), NullObject(), NullObject(), NullObject()]
            )
            changed = True
    if not changed:
        return

    temporary_path = pdf_path.with_name(pdf_path.stem + ".links.tmp.pdf")
    try:
        with temporary_path.open("wb") as target:
            writer.write(target)
        temporary_path.replace(pdf_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


async def export_pdf(html_path: Path, output_path: Path, approved: bool = False) -> Path:
    if not approved:
        raise PermissionError("PDF export requires explicit user approval (--approved).")
    try:
        from bs4 import BeautifulSoup
        from playwright.async_api import async_playwright
        from pypdf import PdfReader, PdfWriter
        from reportlab.lib.colors import HexColor
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.utils import ImageReader
        from reportlab.pdfgen import canvas
    except ImportError as exc:
        raise RuntimeError("Missing PDF dependency. See references/conversion.md.") from exc

    soup = BeautifulSoup(html_path.read_text(encoding="utf-8"), "html.parser")
    meta = {tag.get("name", "")[5:]: tag.get("content", "") for tag in soup.find_all("meta") if tag.get("name", "").startswith("apta:")}
    title = meta.get("title", html_path.stem)
    subtitle = meta.get("subtitle", "Aptaworks Proposal Series")
    status = meta.get("status", "Confidential")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    cover_tmp = output_path.with_name(output_path.stem + ".cover.tmp.pdf")
    content_tmp = output_path.with_name(output_path.stem + ".content.tmp.pdf")
    content_html_tmp = output_path.with_name(output_path.stem + ".content.tmp.html")
    try:
        content_soup = BeautifulSoup(str(soup), "html.parser")
        cover = content_soup.select_one(".cover-page")
        if cover:
            cover.decompose()
        for style in content_soup.find_all("style"):
            if style.string:
                style.string.replace_with(re.sub(r"@page\s*:first\s*\{.*?\}", "", style.string, flags=re.DOTALL | re.IGNORECASE))
        content_section = content_soup.select_one(".content-page")
        if content_section:
            content_section["class"] = ["content-page"]
        override = content_soup.new_tag("style")
        override.string = "@page{size:A4 portrait;margin:30mm 25mm!important}.content-page{min-height:0!important;padding:0!important}"
        content_soup.head.append(override)
        content_html_tmp.write_text(str(content_soup), encoding="utf-8")
        async with async_playwright() as playwright:
            browser = await playwright.chromium.launch(headless=True)
            page = await browser.new_page(viewport={"width": 794, "height": 1123})
            await page.goto(html_path.resolve().as_uri(), wait_until="networkidle")
            await page.emulate_media(media="print")
            await page.pdf(path=str(cover_tmp), format="A4", print_background=True, prefer_css_page_size=True, page_ranges="1")
            await page.goto(content_html_tmp.resolve().as_uri(), wait_until="networkidle")
            await page.emulate_media(media="print")
            await page.pdf(path=str(content_tmp), format="A4", print_background=True, prefer_css_page_size=True)
            await browser.close()
        writer = PdfWriter()
        cover_reader = PdfReader(str(cover_tmp))
        content_reader = PdfReader(str(content_tmp))
        for page in cover_reader.pages:
            writer.add_page(page)
        logo_path = Path(__file__).resolve().parent.parent / "assets" / "logo_aptaworks_wide_navy.png"
        page_width, page_height = A4
        mm = 72 / 25.4
        total_pages = len(cover_reader.pages) + len(content_reader.pages)
        for index, page in enumerate(content_reader.pages, start=len(cover_reader.pages) + 1):
            packet = io.BytesIO()
            overlay = canvas.Canvas(packet, pagesize=A4)
            overlay.setStrokeColor(HexColor("#EF7636"))
            overlay.setLineWidth(0.75)
            overlay.line(25 * mm, page_height - 25 * mm, page_width - 25 * mm, page_height - 25 * mm)
            overlay.setFillColor(HexColor("#17234E"))
            overlay.setFont("Helvetica-Bold", 8)
            overlay.drawString(25 * mm, page_height - 21.5 * mm, status)
            status_width = overlay.stringWidth(status, "Helvetica-Bold", 8)
            overlay.setFillColor(HexColor("#5E6B82"))
            overlay.setFont("Helvetica", 8)
            overlay.drawString(25 * mm + status_width + 3, page_height - 21.5 * mm, f"- {title}")
            logo_width = 31 * mm
            logo_height = logo_width * 63 / 314
            logo_x = page_width - 25 * mm - logo_width
            overlay.drawImage(ImageReader(str(logo_path)), logo_x, page_height - 23 * mm, width=logo_width, height=logo_height, preserveAspectRatio=True, mask="auto")
            label = "Aptaworks Proposal Series"
            label_width = overlay.stringWidth(label, "Helvetica", 8)
            overlay.drawString(logo_x - label_width - 3 * mm, page_height - 21.5 * mm, label)
            overlay.setStrokeColor(HexColor("#C9D2DF"))
            overlay.line(25 * mm, 22 * mm, page_width - 25 * mm, 22 * mm)
            overlay.setFont("Helvetica", 7.5)
            footer_text = subtitle if len(subtitle) <= 90 else subtitle[:87] + "..."
            overlay.drawString(25 * mm, 17.5 * mm, footer_text)
            page_label = f"Page {index} of {total_pages}" if meta.get('language') == 'en' else f"Halaman {index} dari {total_pages}"
            overlay.drawRightString(page_width - 25 * mm, 17.5 * mm, page_label)
            overlay.save()
            packet.seek(0)
            overlay_page = PdfReader(packet).pages[0]
            page.merge_page(overlay_page)
            writer.add_page(page)
        with output_path.open("wb") as target:
            writer.write(target)
        writer.close()
        _repair_internal_links(output_path, html_path)
        return output_path
    finally:
        for temporary in (cover_tmp, content_tmp, content_html_tmp):
            if temporary.exists():
                temporary.unlink()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("html", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--approved", action="store_true", help="Confirm that the user explicitly approved PDF export.")
    args = parser.parse_args()
    output = (args.output or args.html.with_suffix(".pdf")).resolve()
    try:
        result = asyncio.run(export_pdf(args.html.resolve(), output, approved=args.approved))
    except (OSError, PermissionError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"PDF: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
