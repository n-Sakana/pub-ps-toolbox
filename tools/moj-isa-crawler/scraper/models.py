from __future__ import annotations

import dataclasses


@dataclasses.dataclass
class PageRecord:
    order: int
    depth: int
    section: str
    url: str
    status_code: int
    content_type: str
    title: str
    h1: str
    breadcrumb: str
    headings_json: str
    body_text: str
    text_length: int
    table_count: int
    tables_json: str
    link_count: int
    internal_page_link_count: int
    pdf_link_count: int
    fetched_at: str

    def as_dict(self) -> dict[str, object]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class LinkRecord:
    source_page_url: str
    source_page_title: str
    link_text: str
    url: str
    category: str

    def as_dict(self) -> dict[str, object]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class PdfRecord:
    source_page_url: str
    source_page_title: str
    source_section: str
    link_text: str
    pdf_url: str
    filename: str
    downloaded: bool = False
    saved_path: str = ""
    final_url: str = ""
    status_code: int = 0
    content_type: str = ""
    content_length: str = ""
    last_modified: str = ""
    bytes_written: int = 0
    sha256: str = ""
    response_headers_json: str = ""
    first_bytes_hex: str = ""
    post_write_size: int = 0
    post_write_readback_ok: bool = False
    error: str = ""

    def as_dict(self) -> dict[str, object]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class ErrorRecord:
    phase: str
    url: str
    source_page_url: str
    message: str
    exception_type: str = ""
    details: str = ""

    def as_dict(self) -> dict[str, object]:
        return dataclasses.asdict(self)
