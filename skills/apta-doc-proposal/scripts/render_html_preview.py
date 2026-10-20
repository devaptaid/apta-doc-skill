#!/usr/bin/env python3
"""Render a local HTML file to a PNG preview for visual QA."""

from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path


async def render_preview(html_path: Path, output_path: Path) -> Path:
    try:
        from playwright.async_api import async_playwright
    except ImportError as exc:
        raise RuntimeError("Missing dependency 'playwright'. See references/conversion.md.") from exc
    output_path.parent.mkdir(parents=True, exist_ok=True)
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=True)
        page = await browser.new_page(viewport={"width": 1200, "height": 900}, device_scale_factor=1)
        await page.goto(html_path.resolve().as_uri(), wait_until="networkidle")
        await page.screenshot(path=str(output_path), full_page=True)
        await browser.close()
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("html", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    try:
        result = asyncio.run(render_preview(args.html.resolve(), args.output.resolve()))
    except (OSError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"PREVIEW: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
