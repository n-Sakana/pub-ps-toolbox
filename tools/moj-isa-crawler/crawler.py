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
import logging
import sys
import threading
import traceback
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from scraper.analytics import write_graphs
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
    section_from_url,
    tables_json,
    visible_body_text,
)

ROOT = Path(__file__).resolve().parent
DEFAULT_URL_FILE = ROOT / "URL.txt"
DEFAULT_CONFIG_FILE = ROOT / "config.json"


class CrawlError(RuntimeError):
    pass


@dataclass
class PageCrawlResult:
    requested_url: str
    source_page_url: str
    depth: int
    final_url: str = ""
    page: PageRecord | None = None
    links: list[LinkRecord] = field(default_factory=list)
    pdfs: list[PdfRecord] = field(default_factory=list)
    discovered_page_urls: list[str] = field(default_factory=list)
    errors: list[ErrorRecord] = field(default_factory=list)


@dataclass
class PdfTaskResult:
    pdf_url: str
    source_page_url: str
    meta: dict[str, object]
    errors: list[ErrorRecord] = field(default_factory=list)


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
        "default_graph_dir",
        "default_log_file",
        "default_error_log_file",
        "generate_graphs_by_default",
        "log_level",
        "progress_every_pages",
        "progress_every_pdfs",
        "download_pdfs_by_default",
        "probe_pdfs_by_default",
        "timeout",
        "sleep",
        "retries",
        "max_pages",
        "max_depth",
        "max_pdf_downloads",
        "max_page_workers",
        "max_pdf_workers",
        "strict",
    ]
    missing = [key for key in required if key not in config]
    if missing:
        raise CrawlError(f"Config missing required keys: {', '.join(missing)}")
    if not isinstance(config["allowed_prefixes"], list) or not config["allowed_prefixes"]:
        raise CrawlError("Config allowed_prefixes must be a non-empty list")
    for key in ["max_page_workers", "max_pdf_workers"]:
        try:
            value = int(config[key])
        except (TypeError, ValueError) as exc:
            raise CrawlError(f"Config {key} must be an integer >= 1") from exc
        if value < 1:
            raise CrawlError(f"Config {key} must be an integer >= 1")
    if config["max_pdf_downloads"] is not None:
        try:
            max_pdf_downloads = int(config["max_pdf_downloads"])
        except (TypeError, ValueError) as exc:
            raise CrawlError("Config max_pdf_downloads must be null or an integer >= 0") from exc
        if max_pdf_downloads < 0:
            raise CrawlError("Config max_pdf_downloads must be null or an integer >= 0")
    return config


def positive_int(value: str) -> int:
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be >= 1")
    return number


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
    parser.add_argument("--graph-dir", type=Path, help="write diagram PNGs to this directory")
    parser.add_argument("--log-file", type=Path, help="write detailed run log to this text file")
    parser.add_argument("--error-log-file", type=Path, help="write error report to this text file")
    parser.add_argument("--log-level", choices=["DEBUG", "INFO", "WARNING", "ERROR"], help="CLI/file log verbosity")
    parser.add_argument("--progress-every-pages", type=positive_int_or_none, help="log page crawl progress every N pages")
    parser.add_argument("--progress-every-pdfs", type=positive_int_or_none, help="log PDF probe/download progress every N PDFs")
    parser.add_argument("--download-pdfs", action="store_true", help="download discovered PDF files")
    parser.add_argument("--probe-pdfs", action="store_true", help="HEAD request discovered PDFs without saving files")
    parser.add_argument("--max-pages", type=positive_int_or_none)
    parser.add_argument("--max-depth", type=positive_int_or_none)
    parser.add_argument("--max-pdf-downloads", type=positive_int_or_none)
    parser.add_argument("--sleep", type=float)
    parser.add_argument("--timeout", type=int)
    parser.add_argument("--retries", type=int)
    parser.add_argument("--max-page-workers", type=positive_int, help="parallel HTML page fetch workers")
    parser.add_argument("--max-pdf-workers", type=positive_int, help="parallel PDF probe/download workers")
    parser.add_argument("--no-graphs", action="store_true", help="skip PNG graph generation")
    parser.add_argument("--strict", action="store_true", help="fail immediately on page/PDF errors")
    parser.add_argument("--allow-errors", action="store_true", help="write workbook even when some pages/PDFs fail")
    return parser.parse_args(argv)


def setup_logging(*, log_file: Path, log_level: str) -> logging.Logger:
    logger = logging.getLogger("moj_isa_crawler")
    logger.handlers.clear()
    logger.propagate = False
    logger.setLevel(logging.DEBUG)

    numeric_level = getattr(logging, log_level.upper(), logging.INFO)
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setLevel(numeric_level)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    log_file.parent.mkdir(parents=True, exist_ok=True)
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(numeric_level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger


def should_log_progress(current: int, every: int | None) -> bool:
    if every is None or every <= 0:
        return False
    return current == 1 or current % every == 0


def write_error_report(path: Path, errors: list[ErrorRecord]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    now = dt.datetime.now(dt.timezone.utc).isoformat()
    lines = [
        "MOJ ISA crawler error report",
        f"generated_at: {now}",
        f"errors: {len(errors)}",
        "",
    ]
    if not errors:
        lines.append("No errors.")
    else:
        for index, error in enumerate(errors, start=1):
            lines.extend(
                [
                    f"[{index}]",
                    f"phase: {error.phase}",
                    f"url: {error.url}",
                    f"source_page_url: {error.source_page_url}",
                    f"message: {error.message}",
                    f"exception_type: {error.exception_type}",
                    "details:",
                    error.details or "(none)",
                    "",
                ]
            )
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def looks_like_html(content_type: str, content: bytes) -> bool:
    lower_type = content_type.lower()
    if "html" in lower_type:
        return True
    if lower_type and "text/plain" not in lower_type:
        return False
    probe = content[:2048].lower()
    return b"<html" in probe or b"<!doctype html" in probe or b"<head" in probe or b"<body" in probe


def error_from_exception(*, phase: str, url: str, source_page_url: str, exc: BaseException) -> ErrorRecord:
    errno_value = getattr(exc, "errno", None)
    winerror_value = getattr(exc, "winerror", None)
    details = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    if errno_value is not None or winerror_value is not None:
        details = f"errno={errno_value} winerror={winerror_value}\n{details}"
    return ErrorRecord(
        phase=phase,
        url=url,
        source_page_url=source_page_url,
        message=str(exc),
        exception_type=type(exc).__name__,
        details=details,
    )


def error_from_message(*, phase: str, url: str, source_page_url: str, message: str, exception_type: str = "") -> ErrorRecord:
    return ErrorRecord(
        phase=phase,
        url=url,
        source_page_url=source_page_url,
        message=message,
        exception_type=exception_type,
        details=message,
    )


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


def crawl_page_task(
    *,
    url: str,
    depth: int,
    source_page_url: str,
    start_netloc: str,
    allowed_prefixes: list[str],
    same_netloc_only: bool,
    timeout: int,
    sleep: float,
    retries: int,
    logger: logging.Logger,
) -> PageCrawlResult:
    """Fetch and parse one HTML page.

    Each task owns its own Fetcher/requests.Session. requests.Session is not a
    good shared teapot; giving every worker its own one avoids thread-safety
    ambiguity while still parallelizing the network wait.
    """

    worker = threading.current_thread().name
    result = PageCrawlResult(requested_url=url, source_page_url=source_page_url, depth=depth)
    try:
        logger.info("PAGE_FETCH_START worker=%s depth=%s url=%s source=%s", worker, depth, url, source_page_url)
        fetcher = Fetcher(timeout=timeout, sleep=sleep, retries=retries, user_agent=DEFAULT_USER_AGENT)
        fetched = fetcher.get(url)
        result.final_url = fetched.final_url
        if not looks_like_html(fetched.content_type, fetched.content):
            message = (
                f"Non-HTML response: status={fetched.status_code} "
                f"content_type={fetched.content_type!r} bytes={len(fetched.content)}"
            )
            result.errors.append(
                error_from_message(
                    phase="page_non_html",
                    url=fetched.final_url,
                    source_page_url=source_page_url,
                    message=message,
                    exception_type="NonHtmlResponse",
                )
            )
            logger.warning(
                "PAGE_SKIP_NON_HTML worker=%s depth=%s status=%s content_type=%s bytes=%s url=%s source=%s",
                worker,
                depth,
                fetched.status_code,
                fetched.content_type,
                len(fetched.content),
                fetched.final_url,
                source_page_url,
            )
            return result

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
        result.page = PageRecord(
            order=0,
            depth=depth,
            section=section_from_url(fetched.final_url),
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
        for link in page_links:
            result.links.append(
                LinkRecord(
                    source_page_url=fetched.final_url,
                    source_page_title=title,
                    link_text=link.text,
                    url=link.url,
                    category=link.category,
                )
            )
            if link.category == "pdf":
                result.pdfs.append(
                    PdfRecord(
                        source_page_url=fetched.final_url,
                        source_page_title=title,
                        source_section=section_from_url(fetched.final_url),
                        link_text=link.text,
                        pdf_url=link.url,
                        filename=safe_filename_from_url(link.url),
                    )
                )
            elif link.category == "internal_page" and is_allowed_by_prefix(link.url, allowed_prefixes):
                result.discovered_page_urls.append(link.url)
        logger.info(
            "PAGE_FETCH_OK worker=%s depth=%s status=%s page_links=%s pdf_links=%s url=%s title=%s",
            worker,
            depth,
            fetched.status_code,
            result.page.internal_page_link_count,
            result.page.pdf_link_count,
            fetched.final_url,
            title,
        )
        return result
    except Exception as exc:
        result.errors.append(error_from_exception(phase="page", url=url, source_page_url=source_page_url, exc=exc))
        logger.error("PAGE_ERROR worker=%s url=%s source=%s message=%s", worker, url, source_page_url, exc, exc_info=True)
        return result


def handle_pdf_task(
    *,
    pdf_index: int,
    total_pdfs: int,
    pdf_url: str,
    exemplar: PdfRecord,
    mode: str,
    target: Path | None,
    timeout: int,
    sleep: float,
    retries: int,
    log_probe: bool,
    logger: logging.Logger,
) -> PdfTaskResult:
    worker = threading.current_thread().name
    meta: dict[str, object] = {"filename": exemplar.filename}
    try:
        fetcher = Fetcher(timeout=timeout, sleep=sleep, retries=retries, user_agent=DEFAULT_USER_AGENT)
        if mode == "download":
            if target is None:
                raise CrawlError("PDF download task has no target path")
            logger.info(
                "PDF_DOWNLOAD_START worker=%s index=%s total=%s url=%s target=%s",
                worker,
                pdf_index,
                total_pdfs,
                pdf_url,
                target,
            )
            last_logged_mb = -1

            def log_download_progress(bytes_written: int) -> None:
                nonlocal last_logged_mb
                current_mb = bytes_written // (1024 * 1024)
                if current_mb > last_logged_mb:
                    last_logged_mb = current_mb
                    logger.info(
                        "PDF_DOWNLOAD_PROGRESS worker=%s index=%s total=%s bytes=%s target=%s",
                        worker,
                        pdf_index,
                        total_pdfs,
                        bytes_written,
                        target,
                    )

            def log_download_event(event: str, data: dict[str, object]) -> None:
                if event == "response":
                    logger.info(
                        "PDF_RESPONSE worker=%s index=%s total=%s status=%s content_type=%s content_length=%s final_url=%s headers=%s",
                        worker,
                        pdf_index,
                        total_pdfs,
                        data.get("status_code"),
                        data.get("content_type"),
                        data.get("content_length"),
                        data.get("final_url"),
                        data.get("headers"),
                    )
                elif event == "part_verify":
                    logger.info(
                        "PDF_PART_VERIFY worker=%s index=%s total=%s bytes_written=%s part_size=%s first_bytes_hex=%s readback_hex=%s part=%s",
                        worker,
                        pdf_index,
                        total_pdfs,
                        data.get("bytes_written"),
                        data.get("part_size"),
                        data.get("first_bytes_hex"),
                        data.get("readback_hex"),
                        data.get("part_path"),
                    )
                elif event == "final_verify":
                    logger.info(
                        "PDF_FINAL_VERIFY worker=%s index=%s total=%s post_write_size=%s readback_hex=%s readback_ok=%s saved=%s",
                        worker,
                        pdf_index,
                        total_pdfs,
                        data.get("post_write_size"),
                        data.get("readback_hex"),
                        data.get("readback_ok"),
                        data.get("saved_path"),
                    )

            download_result = fetcher.download(pdf_url, target, progress_callback=log_download_progress, event_callback=log_download_event)
            meta.update(
                {
                    "downloaded": True,
                    "saved_path": str(download_result.saved_path),
                    "final_url": download_result.final_url,
                    "status_code": download_result.status_code,
                    "content_type": download_result.content_type,
                    "content_length": download_result.content_length or str(download_result.bytes_written),
                    "last_modified": download_result.last_modified,
                    "bytes_written": download_result.bytes_written,
                    "sha256": download_result.sha256,
                    "response_headers_json": download_result.response_headers_json,
                    "first_bytes_hex": download_result.first_bytes_hex,
                    "post_write_size": download_result.post_write_size,
                    "post_write_readback_ok": download_result.post_write_readback_ok,
                }
            )
            logger.info(
                "PDF_DOWNLOAD_OK worker=%s index=%s total=%s bytes=%s sha256=%s saved=%s",
                worker,
                pdf_index,
                total_pdfs,
                download_result.bytes_written,
                download_result.sha256,
                download_result.saved_path,
            )
        elif mode == "probe":
            if log_probe:
                logger.info("PDF_PROBE_START worker=%s index=%s total=%s url=%s", worker, pdf_index, total_pdfs, pdf_url)
            head_result = fetcher.head(pdf_url)
            meta.update(
                {
                    "downloaded": False,
                    "final_url": head_result.final_url,
                    "status_code": head_result.status_code,
                    "content_type": head_result.content_type,
                    "content_length": head_result.content_length,
                    "last_modified": head_result.last_modified,
                }
            )
            if log_probe:
                logger.info(
                    "PDF_PROBE_OK worker=%s index=%s total=%s status=%s content_type=%s content_length=%s url=%s",
                    worker,
                    pdf_index,
                    total_pdfs,
                    head_result.status_code,
                    head_result.content_type,
                    head_result.content_length,
                    pdf_url,
                )
        else:
            raise CrawlError(f"Unsupported PDF task mode: {mode}")
        return PdfTaskResult(pdf_url=pdf_url, source_page_url=exemplar.source_page_url, meta=meta)
    except Exception as exc:
        meta.update({"error": str(exc)})
        error = error_from_exception(phase="pdf", url=pdf_url, source_page_url=exemplar.source_page_url, exc=exc)
        logger.error(
            "PDF_ERROR worker=%s index=%s total=%s url=%s source=%s message=%s",
            worker,
            pdf_index,
            total_pdfs,
            pdf_url,
            exemplar.source_page_url,
            exc,
            exc_info=True,
        )
        return PdfTaskResult(pdf_url=pdf_url, source_page_url=exemplar.source_page_url, meta=meta, errors=[error])


def crawl(argv: list[str]) -> int:
    args = parse_args(argv)
    config = load_config(args.config)
    start_urls = read_url_file(args.url_file)

    allowed_prefixes = [str(prefix) for prefix in config["allowed_prefixes"]]
    start_netloc = urlparse(start_urls[0]).netloc
    same_netloc_only = bool(config["same_netloc_only"])
    output = args.output or Path(str(config["default_output"]))
    download_dir = args.download_dir or Path(str(config["default_download_dir"]))
    graph_dir = args.graph_dir or Path(str(config["default_graph_dir"]))
    log_file = args.log_file or Path(str(config["default_log_file"]))
    error_log_file = args.error_log_file or Path(str(config["default_error_log_file"]))
    generate_graphs = bool(config["generate_graphs_by_default"]) and not args.no_graphs
    log_level = args.log_level or str(config["log_level"])
    progress_every_pages = args.progress_every_pages if args.progress_every_pages is not None else int(config["progress_every_pages"])
    progress_every_pdfs = args.progress_every_pdfs if args.progress_every_pdfs is not None else int(config["progress_every_pdfs"])
    download_pdfs = bool(config["download_pdfs_by_default"]) or args.download_pdfs
    probe_pdfs = bool(config["probe_pdfs_by_default"]) or args.probe_pdfs or download_pdfs
    timeout = args.timeout if args.timeout is not None else int(config["timeout"])
    sleep = args.sleep if args.sleep is not None else float(config["sleep"])
    retries = args.retries if args.retries is not None else int(config["retries"])
    max_pages = args.max_pages if args.max_pages is not None else config["max_pages"]
    max_depth = args.max_depth if args.max_depth is not None else config["max_depth"]
    max_pdf_downloads = args.max_pdf_downloads if args.max_pdf_downloads is not None else config["max_pdf_downloads"]
    max_page_workers = args.max_page_workers if args.max_page_workers is not None else int(config["max_page_workers"])
    max_pdf_workers = args.max_pdf_workers if args.max_pdf_workers is not None else int(config["max_pdf_workers"])
    if max_pdf_downloads is not None:
        max_pdf_downloads = int(max_pdf_downloads)
        if max_pdf_downloads < 0:
            raise CrawlError("max_pdf_downloads must be >= 0")
    if args.strict and args.allow_errors:
        raise CrawlError("--strict and --allow-errors cannot be used together")
    strict = args.strict or (bool(config["strict"]) and not args.allow_errors)

    logger = setup_logging(log_file=log_file, log_level=log_level)
    logger.info(
        "START start_urls=%s allowed_prefixes=%s output=%s download_pdfs=%s probe_pdfs=%s "
        "download_dir=%s generate_graphs=%s graph_dir=%s max_pages=%s max_depth=%s max_pdf_downloads=%s "
        "page_workers=%s pdf_workers=%s strict=%s log_file=%s error_log_file=%s",
        len(start_urls),
        allowed_prefixes,
        output,
        download_pdfs,
        probe_pdfs,
        download_dir,
        generate_graphs,
        graph_dir,
        max_pages,
        max_depth,
        max_pdf_downloads,
        max_page_workers,
        max_pdf_workers,
        strict,
        log_file,
        error_log_file,
    )

    queue: deque[tuple[str, int, str]] = deque((url, 0, "") for url in start_urls)
    queued: set[str] = set(start_urls)
    visited: set[str] = set()
    completed_page_urls: set[str] = set()
    pages: list[PageRecord] = []
    links: list[LinkRecord] = []
    pdfs: list[PdfRecord] = []
    errors: list[ErrorRecord] = []

    logger.info("PAGE_PHASE_START workers=%s queued=%s", max_page_workers, len(queue))
    while queue:
        if max_pages is not None and len(pages) >= int(max_pages):
            break

        batch: list[tuple[str, int, str]] = []
        remaining_page_slots = None if max_pages is None else max(0, int(max_pages) - len(pages))
        batch_limit = max_page_workers if remaining_page_slots is None else min(max_page_workers, remaining_page_slots)
        if batch_limit <= 0:
            break

        while queue and len(batch) < batch_limit:
            url, depth, source_page_url = queue.popleft()
            if url in visited:
                continue
            if max_depth is not None and depth > int(max_depth):
                continue
            visited.add(url)
            batch.append((url, depth, source_page_url))

        if not batch:
            continue

        logger.info(
            "PAGE_BATCH_START batch_size=%s workers=%s crawled=%s limit=%s queue_remaining=%s",
            len(batch),
            max_page_workers,
            len(pages),
            max_pages if max_pages is not None else "unbounded",
            len(queue),
        )
        with ThreadPoolExecutor(max_workers=max_page_workers, thread_name_prefix="page") as executor:
            futures = [
                executor.submit(
                    crawl_page_task,
                    url=url,
                    depth=depth,
                    source_page_url=source_page_url,
                    start_netloc=start_netloc,
                    allowed_prefixes=allowed_prefixes,
                    same_netloc_only=same_netloc_only,
                    timeout=timeout,
                    sleep=sleep,
                    retries=retries,
                    logger=logger,
                )
                for url, depth, source_page_url in batch
            ]
            for future in as_completed(futures):
                result = future.result()
                if result.final_url:
                    visited.add(result.final_url)
                if result.errors:
                    errors.extend(result.errors)
                    if strict:
                        write_workbook(output, pages=pages, pdfs=pdfs, links=links, errors=errors)
                        write_error_report(error_log_file, errors)
                        first_error = result.errors[0]
                        raise CrawlError(f"Page crawl failed: {first_error.url}: {first_error.message}")
                if result.page is None:
                    continue
                if result.page.url in completed_page_urls:
                    logger.info("PAGE_DUPLICATE_SKIP requested=%s final=%s", result.requested_url, result.page.url)
                    continue

                result.page.order = len(pages) + 1
                pages.append(result.page)
                completed_page_urls.add(result.page.url)
                links.extend(result.links)
                pdfs.extend(result.pdfs)
                if should_log_progress(len(pages), progress_every_pages):
                    logger.info(
                        "PAGE_PROGRESS crawled=%s limit=%s depth=%s status=%s page_links=%s pdf_links=%s queued=%s title=%s url=%s",
                        len(pages),
                        max_pages if max_pages is not None else "unbounded",
                        result.page.depth,
                        result.page.status_code,
                        result.page.internal_page_link_count,
                        result.page.pdf_link_count,
                        len(queue),
                        result.page.title,
                        result.page.url,
                    )

                for discovered_url in result.discovered_page_urls:
                    if discovered_url in visited or discovered_url in queued:
                        continue
                    queued.add(discovered_url)
                    queue.append((discovered_url, result.depth + 1, result.page.url))

        logger.info(
            "PAGE_BATCH_DONE crawled=%s queued=%s visited=%s errors=%s",
            len(pages),
            len(queue),
            len(visited),
            len(errors),
        )

    # Enrich or download unique PDFs, then copy metadata back to every reference.
    pdf_meta: dict[str, dict[str, object]] = {}
    used_paths: set[Path] = set()
    unique_pdf_urls = list(dict.fromkeys(pdf.pdf_url for pdf in pdfs))
    exemplar_by_pdf_url: dict[str, PdfRecord] = {}
    for pdf in pdfs:
        exemplar_by_pdf_url.setdefault(pdf.pdf_url, pdf)
    logger.info(
        "PDF_PHASE_START unique_pdfs=%s references=%s download_pdfs=%s probe_pdfs=%s max_pdf_downloads=%s workers=%s",
        len(unique_pdf_urls),
        len(pdfs),
        download_pdfs,
        probe_pdfs,
        max_pdf_downloads,
        max_pdf_workers,
    )

    pdf_tasks: list[dict[str, object]] = []
    downloads_assigned = 0
    for pdf_index, pdf_url in enumerate(unique_pdf_urls, start=1):
        exemplar = exemplar_by_pdf_url[pdf_url]
        if download_pdfs and (max_pdf_downloads is None or downloads_assigned < max_pdf_downloads):
            target = make_unique_path(download_dir, exemplar.filename, used_paths)
            downloads_assigned += 1
            pdf_tasks.append({"pdf_index": pdf_index, "pdf_url": pdf_url, "exemplar": exemplar, "mode": "download", "target": target})
        elif download_pdfs and max_pdf_downloads is not None:
            pdf_meta[pdf_url] = {"filename": exemplar.filename, "downloaded": False, "error": "skipped_by_max_pdf_downloads"}
            if should_log_progress(pdf_index, progress_every_pdfs):
                logger.info(
                    "PDF_SKIP_PROGRESS index=%s total=%s reason=skipped_by_max_pdf_downloads url=%s",
                    pdf_index,
                    len(unique_pdf_urls),
                    pdf_url,
                )
        elif probe_pdfs:
            pdf_tasks.append({"pdf_index": pdf_index, "pdf_url": pdf_url, "exemplar": exemplar, "mode": "probe", "target": None})
        else:
            pdf_meta[pdf_url] = {"filename": exemplar.filename}
            if should_log_progress(pdf_index, progress_every_pdfs):
                logger.info("PDF_METADATA_ONLY_PROGRESS index=%s total=%s url=%s", pdf_index, len(unique_pdf_urls), pdf_url)

    if pdf_tasks:
        logger.info("PDF_PARALLEL_START tasks=%s workers=%s", len(pdf_tasks), max_pdf_workers)
        completed_pdf_tasks = 0
        with ThreadPoolExecutor(max_workers=max_pdf_workers, thread_name_prefix="pdf") as executor:
            futures = [
                executor.submit(
                    handle_pdf_task,
                    pdf_index=int(task["pdf_index"]),
                    total_pdfs=len(unique_pdf_urls),
                    pdf_url=str(task["pdf_url"]),
                    exemplar=task["exemplar"],  # type: ignore[arg-type]
                    mode=str(task["mode"]),
                    target=task["target"],  # type: ignore[arg-type]
                    timeout=timeout,
                    sleep=sleep,
                    retries=retries,
                    log_probe=should_log_progress(int(task["pdf_index"]), progress_every_pdfs),
                    logger=logger,
                )
                for task in pdf_tasks
            ]
            for future in as_completed(futures):
                result = future.result()
                completed_pdf_tasks += 1
                pdf_meta[result.pdf_url] = result.meta
                if result.errors:
                    errors.extend(result.errors)
                    if strict:
                        write_workbook(output, pages=pages, pdfs=pdfs, links=links, errors=errors)
                        write_error_report(error_log_file, errors)
                        first_error = result.errors[0]
                        raise CrawlError(f"PDF handling failed: {first_error.url}: {first_error.message}")
                if should_log_progress(completed_pdf_tasks, progress_every_pdfs):
                    logger.info(
                        "PDF_PARALLEL_PROGRESS completed=%s tasks=%s unique_pdfs=%s errors=%s",
                        completed_pdf_tasks,
                        len(pdf_tasks),
                        len(unique_pdf_urls),
                        len(errors),
                    )
        logger.info("PDF_PARALLEL_DONE tasks=%s errors=%s", len(pdf_tasks), len(errors))

    for pdf in pdfs:
        meta = pdf_meta.get(pdf.pdf_url, {})
        for key, value in meta.items():
            setattr(pdf, key, value)

    graph_paths: list[Path] = []
    if generate_graphs:
        try:
            logger.info("GRAPH_PHASE_START graph_dir=%s pages=%s links=%s pdf_refs=%s", graph_dir, len(pages), len(links), len(pdfs))
            graph_paths = write_graphs(graph_dir, pages=pages, pdfs=pdfs, links=links, errors=errors)
            for graph_path in graph_paths:
                logger.info("GRAPH_OK path=%s", graph_path)
            logger.info("GRAPH_PHASE_DONE generated=%s graph_dir=%s", len(graph_paths), graph_dir)
        except Exception as exc:
            errors.append(error_from_exception(phase="graph", url=str(graph_dir), source_page_url="", exc=exc))
            logger.error("GRAPH_ERROR graph_dir=%s message=%s", graph_dir, exc, exc_info=True)
            if strict:
                write_workbook(output, pages=pages, pdfs=pdfs, links=links, errors=errors, graph_paths=graph_paths)
                write_error_report(error_log_file, errors)
                raise CrawlError(f"Graph generation failed: {exc}") from exc

    write_workbook(output, pages=pages, pdfs=pdfs, links=links, errors=errors, graph_paths=graph_paths)
    write_error_report(error_log_file, errors)
    downloaded_unique = {p.pdf_url for p in pdfs if p.downloaded}
    logger.info(
        "DONE pages=%s pdf_refs=%s unique_pdfs=%s downloaded_unique_pdfs=%s graphs=%s errors=%s output=%s log_file=%s error_log_file=%s",
        len(pages),
        len(pdfs),
        len({p.pdf_url for p in pdfs}),
        len(downloaded_unique),
        len(graph_paths),
        len(errors),
        output,
        log_file,
        error_log_file,
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
