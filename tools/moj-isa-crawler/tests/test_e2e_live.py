#!/usr/bin/env python3
"""Live E2E test for the ISA/MOJ crawler prototype.

This intentionally accesses https://www.moj.go.jp/isa/. It verifies the workbook
against the current live HTML, not just that a file was created.
"""

from __future__ import annotations

import re
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.parse import urldefrag, urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "crawler.py"
URL_FILE = ROOT / "URL.txt"
CONFIG = ROOT / "config.json"
EXPECTED_MIN_PAGES = 10
EXPECTED_MIN_PDFS = 1
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}


def normalize_space(text: object) -> str:
    return re.sub(r"\s+", " ", str(text or "").replace("\ufeff", "").replace("\u200b", "")).strip()


def fetch_html(url: str) -> str:
    response = requests.get(url, headers=REQUEST_HEADERS, timeout=30)
    response.raise_for_status()
    return response.content.decode("utf-8", errors="replace")


def sheet_records(workbook, sheet_name: str) -> list[dict[str, object]]:
    sheet = workbook[sheet_name]
    rows = sheet.iter_rows(values_only=True)
    headers = [str(value) for value in next(rows)]
    return [dict(zip(headers, row)) for row in rows]


def source_pdf_links(url: str) -> set[str]:
    soup = BeautifulSoup(fetch_html(url), "html.parser")
    links = set()
    for a in soup.find_all("a", href=True):
        linked, _fragment = urldefrag(urljoin(url, a["href"]))
        if ".pdf" in linked.lower():
            links.add(linked)
    return links


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="moj-isa-crawler-e2e-") as tmpdir:
        tmp = Path(tmpdir)
        output = tmp / "moj_isa_crawl.xlsx"
        download_dir = tmp / "pdfs"
        cmd = [
            sys.executable,
            str(SCRIPT),
            "--url-file",
            str(URL_FILE),
            "--config",
            str(CONFIG),
            "--output",
            str(output),
            "--download-dir",
            str(download_dir),
            "--sleep",
            "0",
            "--max-pages",
            "15",
            "--download-pdfs",
            "--max-pdf-downloads",
            "1",
        ]
        completed = subprocess.run(cmd, cwd=ROOT, text=True, capture_output=True)
        print(completed.stdout, end="")
        if completed.stderr:
            print(completed.stderr, file=sys.stderr, end="")
        if completed.returncode != 0:
            return completed.returncode
        if not output.exists() or output.stat().st_size == 0:
            raise AssertionError("Workbook was not created")

        workbook = load_workbook(output, read_only=True, data_only=True)
        for sheet in ["Pages", "PDFs", "Links", "Errors", "Summary"]:
            if sheet not in workbook.sheetnames:
                raise AssertionError(f"Missing sheet: {sheet}")

        pages = sheet_records(workbook, "Pages")
        pdfs = sheet_records(workbook, "PDFs")
        links = sheet_records(workbook, "Links")
        if len(pages) < EXPECTED_MIN_PAGES:
            raise AssertionError(f"Expected at least {EXPECTED_MIN_PAGES} pages, got {len(pages)}")
        if len(pdfs) < EXPECTED_MIN_PDFS:
            raise AssertionError(f"Expected at least {EXPECTED_MIN_PDFS} PDF refs, got {len(pdfs)}")
        if not links:
            raise AssertionError("Links sheet is empty")

        # Page rows are checked against live HTML titles/text.
        for row in pages[:5]:
            url = normalize_space(row["url"])
            live_soup = BeautifulSoup(fetch_html(url), "html.parser")
            live_text = normalize_space(live_soup.get_text(" ", strip=True))
            title = normalize_space(row["title"])
            if title and title[:20] not in live_text:
                raise AssertionError(f"Page title probe not found in live HTML: {url}: {title!r}")
            body = normalize_space(row["body_text"])
            if body:
                probe = body[:30]
                if probe not in live_text:
                    raise AssertionError(f"Body text probe not found in live HTML: {url}: {probe!r}")

        # PDF rows are checked against the live source page's actual hrefs.
        by_source: dict[str, set[str]] = {}
        for row in pdfs:
            source = normalize_space(row["source_page_url"])
            pdf_url = normalize_space(row["pdf_url"])
            if source not in by_source:
                by_source[source] = source_pdf_links(source)
            if pdf_url not in by_source[source]:
                raise AssertionError(f"PDF URL is not present on live source page: {pdf_url}")

        downloaded = [row for row in pdfs if str(row.get("downloaded", "")).lower() == "true"]
        if not downloaded:
            raise AssertionError("No PDF row was marked downloaded")
        downloaded_path = Path(normalize_space(downloaded[0]["saved_path"]))
        sha = normalize_space(downloaded[0]["sha256"])
        if not downloaded_path.exists() or downloaded_path.stat().st_size == 0:
            raise AssertionError(f"Downloaded PDF missing or empty: {downloaded_path}")
        if len(sha) != 64:
            raise AssertionError(f"Invalid sha256: {sha}")

        print(f"E2E OK: pages={len(pages)}, pdf_refs={len(pdfs)}, downloaded={downloaded_path.name}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
