from __future__ import annotations

import dataclasses


@dataclasses.dataclass
class PageRecord:
    order: int
    depth: int
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
    link_text: str
    pdf_url: str
    filename: str
    downloaded: bool = False
    saved_path: str = ""
    content_type: str = ""
    content_length: str = ""
    last_modified: str = ""
    sha256: str = ""
    error: str = ""

    def as_dict(self) -> dict[str, object]:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class ErrorRecord:
    phase: str
    url: str
    source_page_url: str
    message: str

    def as_dict(self) -> dict[str, object]:
        return dataclasses.asdict(self)
