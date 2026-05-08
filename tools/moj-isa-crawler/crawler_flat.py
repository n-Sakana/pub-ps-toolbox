#!/usr/bin/env python3
"""Crawl ISA/MOJ pages and export a flat URL/title/body/links table.

This is a smaller-output variant of crawler.py.  It walks the same allowed
HTML page scope, but skips PDF probing/downloading and graph generation.
"""

from __future__ import annotations

import argparse
import sys
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from crawler import (
    CrawlError,
    DEFAULT_CONFIG_FILE,
    DEFAULT_URL_FILE,
    crawl_page_task,
    load_config,
    positive_int,
    positive_int_or_none,
    read_url_file,
    setup_logging,
    should_log_progress,
    write_error_report,
)
from scraper.exporter import ILLEGAL_XML_RE, excel_safe
from scraper.models import LinkRecord


DEFAULT_OUTPUT = Path("moj_isa_pages_text_links.csv")
DEFAULT_LOG_FILE = Path("logs/moj_isa_pages_text_links.log")
DEFAULT_ERROR_LOG_FILE = Path("logs/moj_isa_pages_text_links_errors.txt")
FLAT_COLUMNS = ["URL", "ページ名", "ページ本文", "含まれるリンク"]


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Crawl ISA/MOJ pages and export URL, page name, body text, and all contained links"
    )
    parser.add_argument("--url-file", type=Path, default=DEFAULT_URL_FILE)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_FILE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="CSV by default; .xlsx writes an Excel workbook")
    parser.add_argument("--log-file", type=Path, default=DEFAULT_LOG_FILE)
    parser.add_argument("--error-log-file", type=Path, default=DEFAULT_ERROR_LOG_FILE)
    parser.add_argument("--log-level", choices=["DEBUG", "INFO", "WARNING", "ERROR"], help="CLI/file log verbosity")
    parser.add_argument("--progress-every-pages", type=positive_int_or_none, help="log page crawl progress every N pages")
    parser.add_argument("--max-pages", type=positive_int_or_none)
    parser.add_argument("--max-depth", type=positive_int_or_none)
    parser.add_argument("--sleep", type=float)
    parser.add_argument("--timeout", type=int)
    parser.add_argument("--retries", type=int)
    parser.add_argument("--max-page-workers", type=positive_int, help="parallel HTML page fetch workers")
    parser.add_argument("--link-separator", default=";", help="separator used in the link column")
    parser.add_argument("--strict", action="store_true", help="fail immediately on page errors")
    parser.add_argument("--allow-errors", action="store_true", help="write output even when some pages fail")
    return parser.parse_args(argv)


def csv_safe(value: object) -> object:
    if isinstance(value, str):
        return ILLEGAL_XML_RE.sub("", value)
    return value


def unique_link_urls(links: list[LinkRecord]) -> list[str]:
    urls: list[str] = []
    seen: set[str] = set()
    for link in links:
        url = link.url.strip()
        if not url or url in seen:
            continue
        seen.add(url)
        urls.append(url)
    return urls


def row_from_result_links(*, url: str, title: str, body_text: str, links: list[LinkRecord], separator: str) -> dict[str, object]:
    return {
        "URL": url,
        "ページ名": title,
        "ページ本文": body_text,
        "含まれるリンク": separator.join(unique_link_urls(links)),
    }


def write_flat_output(output: Path, rows: list[dict[str, object]]) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    frame = pd.DataFrame(rows, columns=FLAT_COLUMNS)
    if output.suffix.lower() in {".xlsx", ".xlsm"}:
        excel_frame = frame.apply(lambda column: column.map(excel_safe))
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excel_frame.to_excel(writer, sheet_name="Pages", index=False)
            worksheet = writer.sheets["Pages"]
            worksheet.freeze_panes = "A2"
            header_fill = PatternFill("solid", fgColor="1F4E78")
            for cell in worksheet[1]:
                cell.font = Font(color="FFFFFF", bold=True)
                cell.fill = header_fill
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for column_cells in worksheet.columns:
                header = str(column_cells[0].value or "")
                width = min(80, max(12, len(header) + 2))
                for cell in column_cells[1:80]:
                    value = cell.value
                    if value is not None:
                        width = min(80, max(width, min(80, len(str(value)) + 2)))
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                worksheet.column_dimensions[column_cells[0].column_letter].width = width
    else:
        csv_frame = frame.apply(lambda column: column.map(csv_safe))
        csv_frame.to_csv(output, index=False, encoding="utf-8-sig")


def crawl_flat(argv: list[str]) -> int:
    args = parse_args(argv)
    config = load_config(args.config)
    start_urls = read_url_file(args.url_file)

    allowed_prefixes = [str(prefix) for prefix in config["allowed_prefixes"]]
    start_netloc = urlparse(start_urls[0]).netloc
    same_netloc_only = bool(config["same_netloc_only"])
    log_level = args.log_level or str(config["log_level"])
    progress_every_pages = (
        args.progress_every_pages if args.progress_every_pages is not None else int(config["progress_every_pages"])
    )
    timeout = args.timeout if args.timeout is not None else int(config["timeout"])
    sleep = args.sleep if args.sleep is not None else float(config["sleep"])
    retries = args.retries if args.retries is not None else int(config["retries"])
    max_pages = args.max_pages if args.max_pages is not None else config["max_pages"]
    max_depth = args.max_depth if args.max_depth is not None else config["max_depth"]
    max_page_workers = args.max_page_workers if args.max_page_workers is not None else int(config["max_page_workers"])
    if args.strict and args.allow_errors:
        raise CrawlError("--strict and --allow-errors cannot be used together")
    strict = args.strict or (bool(config["strict"]) and not args.allow_errors)

    logger = setup_logging(log_file=args.log_file, log_level=log_level)
    logger.info(
        "START_FLAT start_urls=%s allowed_prefixes=%s output=%s max_pages=%s max_depth=%s "
        "page_workers=%s strict=%s link_separator=%r log_file=%s error_log_file=%s",
        len(start_urls),
        allowed_prefixes,
        args.output,
        max_pages,
        max_depth,
        max_page_workers,
        strict,
        args.link_separator,
        args.log_file,
        args.error_log_file,
    )

    queue: deque[tuple[str, int, str]] = deque((url, 0, "") for url in start_urls)
    queued: set[str] = set(start_urls)
    visited: set[str] = set()
    completed_page_urls: set[str] = set()
    rows: list[dict[str, object]] = []
    errors = []

    logger.info("PAGE_PHASE_START workers=%s queued=%s", max_page_workers, len(queue))
    while queue:
        if max_pages is not None and len(rows) >= int(max_pages):
            break

        batch: list[tuple[str, int, str]] = []
        remaining_page_slots = None if max_pages is None else max(0, int(max_pages) - len(rows))
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
            len(rows),
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
                        write_flat_output(args.output, rows)
                        write_error_report(args.error_log_file, errors)
                        first_error = result.errors[0]
                        raise CrawlError(f"Page crawl failed: {first_error.url}: {first_error.message}")
                if result.page is None:
                    continue
                if result.page.url in completed_page_urls:
                    logger.info("PAGE_DUPLICATE_SKIP requested=%s final=%s", result.requested_url, result.page.url)
                    continue

                result.page.order = len(rows) + 1
                completed_page_urls.add(result.page.url)
                rows.append(
                    row_from_result_links(
                        url=result.page.url,
                        title=result.page.title,
                        body_text=result.page.body_text,
                        links=result.links,
                        separator=args.link_separator,
                    )
                )
                if should_log_progress(len(rows), progress_every_pages):
                    logger.info(
                        "PAGE_PROGRESS crawled=%s limit=%s depth=%s status=%s links=%s queued=%s title=%s url=%s",
                        len(rows),
                        max_pages if max_pages is not None else "unbounded",
                        result.page.depth,
                        result.page.status_code,
                        len(unique_link_urls(result.links)),
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
            len(rows),
            len(queue),
            len(visited),
            len(errors),
        )

    write_flat_output(args.output, rows)
    write_error_report(args.error_log_file, errors)
    logger.info(
        "DONE_FLAT pages=%s errors=%s output=%s log_file=%s error_log_file=%s",
        len(rows),
        len(errors),
        args.output,
        args.log_file,
        args.error_log_file,
    )
    if strict and errors:
        return 1
    return 0


def main() -> int:
    try:
        return crawl_flat(sys.argv[1:])
    except CrawlError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
