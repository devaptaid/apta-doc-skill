#!/usr/bin/env python3
"""Compile Aptaworks proposal Markdown to branded HTML."""

from __future__ import annotations

import argparse
import html
import hashlib
import re
import shutil
import sys
from pathlib import Path
from urllib.parse import quote
from urllib.parse import unquote, urlparse


DEFAULTS = {
    "language": "id",
    "title": "Aptaworks Proposal",
    "subtitle": "A focused proposal for measurable delivery.",
    "client": "Client",
    "author": "Aptaworks Team",
    "date": "",
    "version": "1.0",
    "status": "Confidential",
    "display_number": "01",
    "client_logo": "",
}


def require_markdown():
    try:
        import markdown
    except ImportError as exc:
        raise SystemExit("Missing dependency 'markdown'. See references/conversion.md.") from exc
    return markdown


def parse_frontmatter(text: str) -> tuple[dict[str, str], str]:
    metadata = dict(DEFAULTS)
    if not text.startswith("---\n"):
        return metadata, text
    end = text.find("\n---\n", 4)
    if end < 0:
        raise ValueError("Markdown frontmatter starts with '---' but has no closing delimiter.")
    for raw_line in text[4:end].splitlines():
        if not raw_line.strip() or raw_line.lstrip().startswith("#"):
            continue
        if ":" not in raw_line:
            raise ValueError(f"Invalid frontmatter line: {raw_line}")
        key, value = raw_line.split(":", 1)
        key = key.strip()
        if key in metadata:
            metadata[key] = value.strip().strip('"\'')
    return metadata, text[end + 5 :]


def validate_branding(metadata: dict[str, str], body: str) -> None:
    combined = "\n".join(metadata.values()) + "\n" + body
    errors = []
    if re.search(r"\bAptaWorks\b", combined):
        errors.append("Use 'Aptaworks', not 'AptaWorks'.")
    if re.search(r"Aptaworks\s+Consultant", combined, re.IGNORECASE):
        errors.append("Use 'Aptaworks Team', not 'Aptaworks Consultant'.")
    if errors:
        raise ValueError(" ".join(errors))


def local_file_uri(path: Path) -> str:
    return path.resolve().as_uri()


def package_image_sources(content_html: str, source_dir: Path, output_dir: Path) -> str:
    """Bundle local images next to HTML so a downloaded output stays portable."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(content_html, "html.parser")
    for image in soup.find_all("img"):
        src = image.get("src", "")
        parsed = urlparse(src)
        if parsed.scheme in {"https", "http", "data"}:
            continue
        if parsed.scheme == "file":
            raw = unquote(parsed.path)
            if re.match(r"^/[A-Za-z]:/", raw):
                raw = raw[1:]
            source = Path(raw)
        else:
            source = source_dir / unquote(src)
        if not source.is_file():
            raise ValueError(f"Image not found: {source}")
        digest = hashlib.sha256(source.read_bytes()).hexdigest()[:12]
        target = output_dir / "assets" / f"{digest}-{source.name}"
        target.parent.mkdir(parents=True, exist_ok=True)
        if source.resolve() != target.resolve():
            shutil.copy2(source, target)
        image["src"] = quote(f"assets/{target.name}")
    return str(soup)


def normalize_image_sources(content_html: str, source_dir: Path) -> str:
    """Convert local relative image paths to file URIs for portable rendering."""
    try:
        from bs4 import BeautifulSoup
    except ImportError as exc:
        raise SystemExit("Missing dependency 'beautifulsoup4'. See references/conversion.md.") from exc
    soup = BeautifulSoup(content_html, "html.parser")
    for image in soup.find_all("img"):
        src = image.get("src", "")
        if not src or re.match(r"^(?:https?:|data:|file:)", src, re.IGNORECASE):
            continue
        candidate = (source_dir / Path(src.replace("%20", " "))).resolve()
        image["src"] = local_file_uri(candidate)
    return str(soup)


_MANUAL_HEADING_NUMBER = re.compile(r"^\s*\d+(?:\.\d+)*\.?\s+")


def _clean_heading_text(text: str) -> str:
    """Remove a leading manual section number before the compiler adds one."""
    normalized = re.sub(r"\s+", " ", text).strip()
    return _MANUAL_HEADING_NUMBER.sub("", normalized).strip()


def _slugify_heading(text: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return slug or "section"


def add_numbering_and_toc(content_html: str, title: str = "DAFTAR ISI") -> str:
    """Number h1-h3 headings and prepend a linked Table of Contents page.

    The source Markdown remains semantic: authors write ``# Heading`` rather
    than ``# 1. Heading``. Existing numeric prefixes are stripped so repeated
    compilation never produces ``1. 1. Heading``.
    """
    try:
        from bs4 import BeautifulSoup
    except ImportError as exc:
        raise SystemExit("Missing dependency 'beautifulsoup4'. See references/conversion.md.") from exc

    soup = BeautifulSoup(content_html, "html.parser")
    counters = [0, 0, 0]
    used_ids: set[str] = set()
    toc_items: list[dict[str, str | int]] = []

    for heading in soup.find_all(["h1", "h2", "h3"]):
        raw_text = heading.get_text(" ", strip=True)
        if not raw_text:
            continue

        level = int(heading.name[1])
        cleaned_text = _clean_heading_text(raw_text)
        if level == 1:
            counters[0] += 1
            counters[1:] = [0, 0]
        elif level == 2:
            if counters[0] == 0:
                counters[0] = 1
            counters[1] += 1
            counters[2] = 0
        else:
            if counters[0] == 0:
                counters[0] = 1
            if counters[1] == 0:
                counters[1] = 1
            counters[2] += 1

        number = ".".join(str(value) for value in counters[:level]) + "."
        display = f"{number} {cleaned_text}"
        section_id = _slugify_heading(cleaned_text)
        if section_id in used_ids:
            suffix = 2
            while f"{section_id}-{suffix}" in used_ids:
                suffix += 1
            section_id = f"{section_id}-{suffix}"
        used_ids.add(section_id)

        heading["id"] = section_id
        heading.clear()
        heading.string = display
        toc_items.append({
            "level": level,
            "number": number,
            "text": cleaned_text,
            "display": display,
            "id": section_id,
        })

    toc_page = soup.new_tag("div", attrs={"class": "toc-page"})
    toc_title = soup.new_tag("h1")
    toc_title.string = title
    toc_page.append(toc_title)
    entries = soup.new_tag("div", attrs={"class": "toc-entries"})
    for item in toc_items:
        entry = soup.new_tag(
            "div",
            attrs={
                "class": f"toc-entry toc-level-{item['level']}",
                "data-toc-level": str(item["level"]),
                "data-toc-number": str(item["number"]),
                "data-toc-text": str(item["text"]),
                "data-toc-display": str(item["display"]),
                "data-toc-id": str(item["id"]),
            },
        )
        link = soup.new_tag("a", href=f"#{item['id']}")
        link.string = str(item["display"])
        entry.append(link)
        entry.append(soup.new_tag("span", attrs={"class": "toc-dots"}))
        entry.append(
            soup.new_tag(
                "span",
                attrs={
                    "class": "toc-page-number",
                    "data-toc-page-key": str(item["display"]),
                },
            )
        )
        entries.append(entry)
    toc_page.append(entries)
    soup.insert(0, toc_page)
    return str(soup)


def apply_toc_page_numbers(html_path: Path, page_map: dict[str, int]) -> Path:
    """Fill the PDF/HTML TOC page-number cells after the first PDF pass."""
    try:
        from bs4 import BeautifulSoup
    except ImportError as exc:
        raise SystemExit("Missing dependency 'beautifulsoup4'. See references/conversion.md.") from exc

    soup = BeautifulSoup(html_path.read_text(encoding="utf-8"), "html.parser")
    for entry in soup.select(".toc-entry"):
        display = entry.get("data-toc-display", "")
        page_number = page_map.get(display)
        page_cell = entry.select_one(".toc-page-number")
        if page_cell is None:
            page_cell = soup.new_tag("span", attrs={"class": "toc-page-number"})
            entry.append(page_cell)
        page_cell.string = str(page_number) if page_number is not None else ""
    html_path.write_text(str(soup), encoding="utf-8")
    return html_path


def compile_markdown(md_path: Path, output_path: Path, template_path: Path, logo_path: Path) -> Path:
    text = md_path.read_text(encoding="utf-8")
    metadata, body = parse_frontmatter(text)
    if metadata['language'] not in {'en', 'id'}:
        raise ValueError("language must be 'en' or 'id'.")
    validate_branding(metadata, body)
    markdown = require_markdown()
    content_html = markdown.markdown(body, extensions=["extra", "fenced_code", "tables", "sane_lists"])
    content_html = add_numbering_and_toc(content_html, "CONTENTS" if metadata['language'] == 'en' else "DAFTAR ISI")
    content_html = normalize_image_sources(content_html, md_path.parent)
    template = template_path.read_text(encoding="utf-8")
    values = {key: html.escape(str(value), quote=True) for key, value in metadata.items()}
    values['author_label'] = 'Author' if metadata['language'] == 'en' else 'Penulis'
    values['date_label'] = 'Date' if metadata['language'] == 'en' else 'Tanggal'
    client_logo = metadata.get("client_logo", "").strip()
    if client_logo:
        client_logo_path = (md_path.parent / Path(client_logo.replace("%20", " "))).resolve()
        if not client_logo_path.exists():
            raise ValueError(f"Client logo not found: {client_logo_path}")
        values["client_mark"] = f'<img src="{local_file_uri(client_logo_path)}" alt="{html.escape(metadata["client"], quote=True)}">'
    else:
        values["client_mark"] = f'<span class="client-name">{html.escape(metadata["client"])}</span>'
    values["content_html"] = content_html
    values["logo_uri"] = local_file_uri(logo_path)
    for key, value in values.items():
        template = template.replace("{{" + key + "}}", value)
    unresolved = sorted(set(re.findall(r"\{\{([a-z_]+)\}\}", template)))
    if unresolved:
        raise ValueError(f"Unresolved template fields: {', '.join(unresolved)}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    template = package_image_sources(template, md_path.parent, output_path.parent)
    output_path.write_text(template, encoding="utf-8")
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("markdown", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--logo", type=Path)
    args = parser.parse_args()
    skill_dir = Path(__file__).resolve().parent.parent
    output = args.output or args.markdown.with_suffix(".html")
    template = args.template or skill_dir / "assets" / "proposal-template.html"
    logo = args.logo or skill_dir / "assets" / "logo_aptaworks_wide.svg"
    try:
        result = compile_markdown(args.markdown.resolve(), output.resolve(), template.resolve(), logo.resolve())
    except (OSError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"HTML: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
