#!/usr/bin/env python3
"""Crawl the ISA/MOJ website and export page/PDF inventory to Excel.

Default run is intentionally metadata-only:

    python crawler.py

Use --download-pdfs when binary PDF files should also be saved locally.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from collections import deque
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from scraper.exporter import write_workbook
from scraper.fetcher import DEFAULT_USER_AGENT, Fetcher
from scraper.models import ErrorRecord, LinkRecord, PageRecord, PdfRecord
from scraper.parser import (
    breadcrumb_text,
    extract_links,
    h1_text,
    headings_json,
    is_allowed_by_prefix,
    page_title,
    parse_html,
    safe_filename_from_url,
    tables_json,
    visible_body_text,
)

ROOT = Path(__file__).resolve().parent
DEFAULT_URL_FILE = ROOT / "URL.txt"
DEFAULT_CONFIG_FILE = ROOT / "config.json"


class CrawlError(RuntimeError):
    pass


def read_url_file(path: Path) -> list[str]:
    if not path.exists():
        raise CrawlError(f"URL file does not exist: {path}")
    urls = []
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        value = line.strip()
        if not value or value.startswith("#"):
            continue
        parsed = urlparse(value)
        if parsed.scheme not in {"http", "https"}:
            raise CrawlError(f"URL.txt contains non-http URL: {value}")
        urls.append(value)
    if not urls:
        raise CrawlError(f"URL file has no crawl start URL: {path}")
    return urls


def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise CrawlError(f"Config file does not exist: {path}")
    try:
        config = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise CrawlError(f"Invalid JSON config: {path}: {exc}") from exc
    required = [
        "allowed_prefixes",
        "same_netloc_only",
        "default_output",
        "default_download_dir",
        "download_pdfs_by_default",
        "probe_pdfs_by_default",
        "timeout",
        "sleep",
        "retries",
        "max_pages",
        "max_depth",
        "strict",
    ]
    missing = [key for key in required if key not in config]
    if missing:
        raise CrawlError(f"Config missing required keys: {', '.join(missing)}")
    if not isinstance(config["allowed_prefixes"], list) or not config["allowed_prefixes"]:
        raise CrawlError("Config allowed_prefixes must be a non-empty list")
    return config


def positive_int_or_none(value: str | None) -> int | None:
    if value is None:
        return None
    number = int(value)
    if number < 0:
        raise argparse.ArgumentTypeError("must be >= 0")
    return number


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Crawl ISA/MOJ pages and export page/PDF inventory to Excel")
    parser.add_argument("--url-file", type=Path, default=DEFAULT_URL_FILE)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_FILE)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--download-dir", type=Path)
    parser.add_argument("--download-pdfs", action="store_true", help="download discovered PDF files")
    parser.add_argument("--probe-pdfs", action="store_true", help="HEAD request discovered PDFs without saving files")
    parser.add_argument("--max-pages", type=positive_int_or_none)
    parser.add_argument("--max-depth", type=positive_int_or_none)
    parser.add_argument("--max-pdf-downloads", type=positive_int_or_none)
    parser.add_argument("--sleep", type=float)
    parser.add_argument("--timeout", type=int)
    parser.add_argument("--retries", type=int)
    parser.add_argument("--allow-errors", action="store_true", help="write workbook even when some pages/PDFs fail")
    return parser.parse_args(argv)


def make_unique_path(download_dir: Path, filename: str, used: set[Path]) -> Path:
    base = download_dir / filename
    candidate = base
    stem = base.stem
    suffix = base.suffix or ".pdf"
    index = 2
    while candidate in used or candidate.exists():
        candidate = download_dir / f"{stem}_{index}{suffix}"
        index += 1
    used.add(candidate)
    return candidate


def crawl(argv: list[str]) -> int:
    args = parse_args(argv)
    config = load_config(args.config)
    start_urls = read_url_file(args.url_file)

    allowed_prefixes = [str(prefix) for prefix in config["allowed_prefixes"]]
    start_netloc = urlparse(start_urls[0]).netloc
    same_netloc_only = bool(config["same_netloc_only"])
    output = args.output or Path(str(config["default_output"]))
    download_dir = args.download_dir or Path(str(config["default_download_dir"]))
    download_pdfs = bool(config["download_pdfs_by_default"]) or args.download_pdfs
    probe_pdfs = bool(config["probe_pdfs_by_default"]) or args.probe_pdfs or download_pdfs
    timeout = args.timeout if args.timeout is not None else int(config["timeout"])
    sleep = args.sleep if args.sleep is not None else float(config["sleep"])
    retries = args.retries if args.retries is not None else int(config["retries"])
    max_pages = args.max_pages if args.max_pages is not None else config["max_pages"]
    max_depth = args.max_depth if args.max_depth is not None else config["max_depth"]
    strict = bool(config["strict"]) and not args.allow_errors

    fetcher = Fetcher(timeout=timeout, sleep=sleep, retries=retries, user_agent=DEFAULT_USER_AGENT)
    queue: deque[tuple[str, int]] = deque((url, 0) for url in start_urls)
    queued = set(start_urls)
    visited: set[str] = set()
    pages: list[PageRecord] = []
    links: list[LinkRecord] = []
    pdfs: list[PdfRecord] = []
    errors: list[ErrorRecord] = []

    while queue:
        if max_pages is not None and len(pages) >= int(max_pages):
            break
        url, depth = queue.popleft()
        if url in visited:
            continue
        if max_depth is not None and depth > int(max_depth):
            continue
        visited.add(url)
        try:
            fetched = fetcher.get(url)
            soup = parse_html(fetched.text())
            title = page_title(soup)
            page_links = extract_links(
                soup,
                fetched.final_url,
                start_netloc=start_netloc,
                allowed_prefixes=allowed_prefixes,
                same_netloc_only=same_netloc_only,
            )
            body_text, text_length = visible_body_text(soup)
            table_data, table_count = tables_json(soup)
            page = PageRecord(
                order=len(pages) + 1,
                depth=depth,
                url=fetched.final_url,
                status_code=fetched.status_code,
                content_type=fetched.content_type,
                title=title,
                h1=h1_text(soup),
                breadcrumb=breadcrumb_text(soup),
                headings_json=headings_json(soup),
                body_text=body_text,
                text_length=text_length,
                table_count=table_count,
                tables_json=table_data,
                link_count=len(page_links),
                internal_page_link_count=len([link for link in page_links if link.category == "internal_page"]),
                pdf_link_count=len([link for link in page_links if link.category == "pdf"]),
                fetched_at=dt.datetime.now(dt.timezone.utc).isoformat(),
            )
            pages.append(page)

            for link in page_links:
                links.append(
                    LinkRecord(
                        source_page_url=fetched.final_url,
                        source_page_title=title,
                        link_text=link.text,
                        url=link.url,
                        category=link.category,
                    )
                )
                if link.category == "pdf":
                    pdfs.append(
                        PdfRecord(
                            source_page_url=fetched.final_url,
                            source_page_title=title,
                            link_text=link.text,
                            pdf_url=link.url,
                            filename=safe_filename_from_url(link.url),
                        )
                    )
                elif link.category == "internal_page" and link.url not in visited and link.url not in queued:
                    if is_allowed_by_prefix(link.url, allowed_prefixes):
                        queued.add(link.url)
                        queue.append((link.url, depth + 1))
        except Exception as exc:
            errors.append(ErrorRecord(phase="page", url=url, source_page_url="", message=str(exc)))
            if strict:
                write_workbook(output, pages=pages, pdfs=pdfs, links=links, errors=errors)
                raise CrawlError(f"Page crawl failed: {url}: {exc}") from exc

    # Enrich or download unique PDFs, then copy metadata back to every reference.
    pdf_meta: dict[str, dict[str, object]] = {}
    used_paths: set[Path] = set()
    downloaded_count = 0
    for pdf_url in dict.fromkeys(pdf.pdf_url for pdf in pdfs):
        exemplar = next(pdf for pdf in pdfs if pdf.pdf_url == pdf_url)
        meta: dict[str, object] = {"filename": exemplar.filename}
        try:
            if download_pdfs and (args.max_pdf_downloads is None or downloaded_count < args.max_pdf_downloads):
                target = make_unique_path(download_dir, exemplar.filename, used_paths)
                result = fetcher.download(pdf_url, target)
                downloaded_count += 1
                meta.update(
                    {
                        "downloaded": True,
                        "saved_path": str(result.saved_path),
                        "content_type": result.content_type,
                        "content_length": result.content_length or str(result.bytes_written),
                        "last_modified": result.last_modified,
                        "sha256": result.sha256,
                    }
                )
            elif download_pdfs and args.max_pdf_downloads is not None:
                meta.update({"downloaded": False, "error": "skipped_by_max_pdf_downloads"})
            elif probe_pdfs:
                result = fetcher.head(pdf_url)
                meta.update(
                    {
                        "downloaded": False,
                        "content_type": result.content_type,
                        "content_length": result.content_length,
                        "last_modified": result.last_modified,
                    }
                )
        except Exception as exc:
            meta.update({"error": str(exc)})
            errors.append(ErrorRecord(phase="pdf", url=pdf_url, source_page_url=exemplar.source_page_url, message=str(exc)))
            if strict:
                write_workbook(output, pages=pages, pdfs=pdfs, links=links, errors=errors)
                raise CrawlError(f"PDF handling failed: {pdf_url}: {exc}") from exc
        pdf_meta[pdf_url] = meta

    for pdf in pdfs:
        meta = pdf_meta.get(pdf.pdf_url, {})
        for key, value in meta.items():
            setattr(pdf, key, value)

    write_workbook(output, pages=pages, pdfs=pdfs, links=links, errors=errors)
    downloaded_unique = {p.pdf_url for p in pdfs if p.downloaded}
    print(
        f"OK: pages={len(pages)}, pdf_refs={len(pdfs)}, unique_pdfs={len({p.pdf_url for p in pdfs})}, "
        f"downloaded_unique_pdfs={len(downloaded_unique)}, errors={len(errors)}, output={output}"
    )
    if strict and errors:
        return 1
    return 0


def main() -> int:
    try:
        return crawl(sys.argv[1:])
    except CrawlError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
